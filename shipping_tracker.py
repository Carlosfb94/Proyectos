# -*- coding: utf-8 -*-
"""Tools for extracting and updating shipment information.

This module consolidates the logic to parse PDFs/Excel files from
various carriers and obtain their shipping status via web scraping.
It was adapted from a larger script with several improvements:

* Use dataclasses for readability.
* Provide utility functions for repeated tasks.
* Wrap webdriver usage with a context manager.
* Allow configuration via environment variables.
"""

from __future__ import annotations

import os
import re
import time
from dataclasses import dataclass
from datetime import datetime
from typing import Iterable, List, Optional

import pandas as pd
import pdfplumber
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class Shipment:
    """Simple container for shipping information."""

    carrier: str
    tracking_number: str
    consignee: str
    company: str
    reference: str = ""
    status: str = ""


# ---------------------------------------------------------------------------
# Parsing utilities
# ---------------------------------------------------------------------------


def _load_table_from_excel(path: str, headers: Iterable[str]) -> Optional[pd.DataFrame]:
    """Return a table whose first row matches ``headers`` (case insensitive)."""

    df = pd.read_excel(path, header=None)
    headers_up = [h.upper() for h in headers]
    for idx, row in df.iterrows():
        row_values = [str(v).strip().upper() for v in row]
        if all(h in row_values for h in headers_up):
            table = df.iloc[idx + 1 :].copy()
            table.columns = row_values
            return table.reset_index(drop=True)
    return None


def extract_starken_excel(path: str) -> List[Shipment]:
    """Parse Starken shipments from an Excel file."""

    header_sets = [
        ["ORDEN DE TRANSPORTE", "DESTINATARIO"],
        ["ORDEN TRANSPORTE", "DESTINATARIO"],
        ["NUMERO DE SEGUIMIENTO", "DESTINATARIO"],
    ]
    for headers in header_sets:
        df = _load_table_from_excel(path, headers)
        if df is None:
            continue
        col_order = next(c for c in df.columns if "ORDEN" in c or "NUMERO" in c)
        col_dest = next(c for c in df.columns if "DESTINATARIO" in c)
        return [
            Shipment(
                "Starken",
                str(row.get(col_order, "")).strip(),
                str(row.get(col_dest, "")).strip(),
                "Starken",
            )
            for _, row in df.iterrows()
            if str(row.get(col_order, "")).strip().lower() not in {"", "nan"}
        ]
    print(f"❌ No se encontraron encabezados válidos en: {path}")
    return []


def extract_fedex_pdf(path: str) -> List[Shipment]:
    """Parse FedEx shipments from a PDF."""

    shipments: List[Shipment] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            for row in table[1:]:
                if len(row) >= 3 and row[0] and row[2]:
                    tracking = row[0].strip()
                    consignee = row[2].strip().split("\n")[0]
                    if tracking.isdigit() and len(tracking) >= 8:
                        shipments.append(
                            Shipment("FedEx", tracking, consignee, "FedEx")
                        )
    return shipments


def _reference_ok(text: str) -> bool:
    codes = re.findall(r"F-(\d+)", text)
    return any(c.startswith("36") for c in codes)


def extract_correos_chile_pdf(path: str) -> List[Shipment]:
    """Parse Correos de Chile manifests from a PDF."""

    shipments: List[Shipment] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            headers = [c.strip().upper() for c in table[0]]
            try:
                idx_dest = headers.index("DESTINATARIO")
                idx_ref = headers.index("REFERENCIA")
                idx_track = headers.index("SEGUIMIENTO")
            except ValueError:
                continue
            for row in table[1:]:
                if len(row) <= max(idx_dest, idx_ref, idx_track):
                    continue
                dest = row[idx_dest].strip()
                ref = row[idx_ref].strip()
                track = row[idx_track].strip()
                if _reference_ok(ref) and track.isdigit():
                    shipments.append(
                        Shipment("Correos de Chile", track, dest, "Correos de Chile", ref)
                    )
    return shipments


def extract_cruz_del_sur_excel(path: str) -> List[Shipment]:
    """Parse Cruz del Sur shipments from an Excel file."""

    df = _load_table_from_excel(path, ["ORDEN TRANSPORTE", "DESTINATARIO"])
    if df is None:
        print(f"❌ No se encontraron encabezados válidos en: {path}")
        return []
    return [
        Shipment(
            "Cruz del Sur",
            str(row.get("ORDEN TRANSPORTE", "")).strip(),
            str(row.get("DESTINATARIO", "")).strip(),
            "Cruz del Sur",
        )
        for _, row in df.iterrows()
        if str(row.get("ORDEN TRANSPORTE", "")).strip().lower() not in {"", "nan"}
    ]


# ---------------------------------------------------------------------------
# Status lookup helpers
# ---------------------------------------------------------------------------


def _simple_soup_get(url: str, timeout: int = 15) -> BeautifulSoup:
    r = requests.get(url, timeout=timeout)
    r.raise_for_status()
    return BeautifulSoup(r.text, "html.parser")


def status_fedex(tracking_number: str) -> str:
    url = f"https://clsclweb.tntchile.cl/txapgw/tracking.asp?boleto={tracking_number}"
    try:
        soup = _simple_soup_get(url)
        text = soup.get_text(" ", strip=True)
        if "Situacion:" in text or "Situación:" in text:
            match = re.search(r"Situaci(?:o|ó)n:\s*(.*)", text)
            if match:
                return match.group(1).strip()
        if "ENTREGADA" in text.upper():
            return "Entregada"
    except Exception as exc:  # pragma: no cover - network
        return f"Error: {exc}"
    return "No disponible"


def status_correos_chile(tracking_number: str) -> str:
    url = f"https://www.correos.cl/web/guest/seguimiento-en-linea?numero={tracking_number}"
    try:
        soup = _simple_soup_get(url, timeout=20)
        estado = soup.find("span", {"class": "estado"})
        if estado and estado.text.strip():
            return estado.text.strip()
        text = soup.get_text(" ", strip=True).upper()
        if "ENTREGADO" in text:
            return "Entregado"
        if "NO REGISTRA INFORMACI" in text:
            return "No registra información"
    except Exception as exc:  # pragma: no cover - network
        return f"Error: {exc}"
    return "En tránsito o no disponible"


# ---------------------------------------------------------------------------
# Selenium helpers
# ---------------------------------------------------------------------------


class Chrome:
    """Context manager for a headless Chrome driver."""

    def __init__(self) -> None:
        opts = Options()
        opts.add_argument("--headless=new")
        opts.add_argument("--window-size=1920,1080")
        self.driver = webdriver.Chrome(options=opts)

    def __enter__(self) -> webdriver.Chrome:
        return self.driver

    def __exit__(self, exc_type, exc, tb) -> None:
        self.driver.quit()


def status_starken(tracking_number: str) -> str:
    url = f"https://www.starken.cl/seguimiento?codigo={tracking_number}"
    try:
        with Chrome() as driver:
            driver.get(url)
            time.sleep(3)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)

            posibles = [
                "El envío ya fue entregado",
                "Entregado con fecha",
                "En tránsito",
                "En reparto",
                "En sucursal destino",
                "Recibido por Starken",
                "Solicitud de Envío Creado",
            ]
            estado_text = None
            fecha_text = None
            wait = WebDriverWait(driver, 20)
            for palabra in posibles:
                try:
                    elem = wait.until(
                        EC.presence_of_element_located(
                            (By.XPATH, f"//*[contains(text(),'{palabra}')]")
                        )
                    )
                    estado_text = elem.text.strip()
                    break
                except Exception:
                    continue
            if not estado_text:
                return "No se detectó estado claro."
            try:
                fecha_elem = driver.find_element(By.XPATH, "//*[contains(text(),'Entregado con fecha')]")
                fecha_text = fecha_elem.text.strip()
            except Exception:
                pass
            if fecha_text:
                return f"{estado_text} - {fecha_text.replace('Entregado con fecha ', '')}"
            return estado_text
    except Exception as exc:  # pragma: no cover - network
        return f"Error Selenium: {exc}"


# ---------------------------------------------------------------------------
# Utils for Cruz del Sur (simplified)
# ---------------------------------------------------------------------------


def _parse_date_lines(text: str) -> List[tuple[datetime, str, str]]:
    results = []
    for line in text.splitlines():
        m = re.match(r"(\d{2}/\d{2}/\d{4} \d{2}:\d{2})\s+(.*)", line)
        if m:
            try:
                dt = datetime.strptime(m.group(1), "%d/%m/%Y %H:%M")
                results.append((dt, line.strip(), m.group(2).strip()))
            except ValueError:
                pass
    return results


def consulta_cruz_del_sur(tracking_number: str, *, max_tries: int = 5) -> Optional[str]:
    """Query Cruz del Sur tracking. Requires a captcha bypass."""

    # API key is read from an environment variable so secrets are not hardcoded
    api_key = os.environ.get("API_KEY_2CAPTCHA")
    if not api_key:
        print("API key de 2Captcha no configurada (API_KEY_2CAPTCHA).")
        return None

    for attempt in range(1, max_tries + 1):
        print(f"Consultando Cruz del Sur para {tracking_number} (intento {attempt})...")
        try:
            with Chrome() as driver:
                driver.get("https://www.cruzdelsurcarga.cl/seguimiento/")
                time.sleep(2)
                input_nro = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "nrodoc"))
                )
                input_nro.clear()
                input_nro.send_keys(tracking_number)
                driver.save_screenshot("screenshot.png")
                from PIL import Image  # lazy import

                im = Image.open("screenshot.png")
                captcha_im = im.crop((563, 409, 701, 471))
                captcha_im.save("captcha_crop.png")
                with open("captcha_crop.png", "rb") as f:
                    r = requests.post(
                        "http://2captcha.com/in.php",
                        files={"file": f},
                        data={"key": api_key, "method": "post"},
                        timeout=20,
                    )
                if "OK|" not in r.text:
                    print("Error enviando captcha:", r.text)
                    continue
                captcha_id = r.text.split("|")[1]
                captcha_result = None
                for _ in range(15):
                    time.sleep(5)
                    res = requests.get(
                        f"http://2captcha.com/res.php?key={api_key}&action=get&id={captcha_id}",
                        timeout=20,
                    )
                    if res.text == "CAPCHA_NOT_READY":
                        continue
                    if "OK|" in res.text:
                        captcha_result = res.text.split("|")[1]
                        break
                    print("Error captcha:", res.text)
                    break
                if not captcha_result:
                    continue
                input_captcha = driver.find_element(By.ID, "captcha")
                input_captcha.clear()
                input_captcha.send_keys(captcha_result)
                for elem in driver.find_elements(By.TAG_NAME, "input"):
                    if "CONSULTAR" in (elem.get_attribute("value") or "").upper():
                        elem.click()
                        break
                else:
                    print("Botón CONSULTAR no encontrado")
                    continue
                time.sleep(6)
                tables = driver.find_elements(By.TAG_NAME, "table")
                all_dates = []
                for table in tables:
                    all_dates.extend(_parse_date_lines(table.text))
                if all_dates:
                    all_dates.sort(key=lambda x: x[0], reverse=True)
                    dt, _line, status = all_dates[0]
                    return f"{status} [{dt:%d/%m/%Y %H:%M}]"
        except Exception as exc:
            print("Fallo en la consulta:", exc)
        time.sleep(3)
    print("Falló la consulta Cruz del Sur después de varios intentos.")
    return None


# ---------------------------------------------------------------------------
# High level workflow helpers
# ---------------------------------------------------------------------------


def update_status(shipment: Shipment, cruz_update: Optional[tuple[str, str]] = None) -> Shipment:
    """Update shipment status in place and return it."""

    carrier = shipment.carrier.lower()
    tracking = shipment.tracking_number
    if carrier == "fedex":
        shipment.status = status_fedex(tracking)
    elif carrier == "correos de chile":
        shipment.status = status_correos_chile(tracking)
    elif carrier == "starken":
        shipment.status = status_starken(tracking)
    elif carrier == "cruz del sur":
        if cruz_update and tracking == cruz_update[0]:
            shipment.status = cruz_update[1]
        else:
            shipment.status = shipment.status or "Requiere consulta manual"
    else:
        shipment.status = shipment.status or "Sin definir"
    return shipment


# ---------------------------------------------------------------------------
# Example CLI
# ---------------------------------------------------------------------------


def main() -> None:  # pragma: no cover - CLI helper
    """Simple CLI for processing a directory of shipping files."""

    import argparse

    parser = argparse.ArgumentParser(description="Process shipping manifests")
    parser.add_argument("directory", help="Folder with PDF/XLSX files")
    parser.add_argument(
        "--excel",
        default="envios.xlsx",
        help="Excel file where results are stored (default: envios.xlsx)",
    )
    args = parser.parse_args()

    excel_path = os.path.abspath(args.excel)
    if not os.path.exists(excel_path):
        pd.DataFrame(columns=[
            "Tipo",
            "Numero de Seguimiento/Orden",
            "Consignatario/Destinatario",
            "Compañía de Envío",
            "Referencia",
            "Estado",
        ]).to_excel(excel_path, index=False)
        print(f"Creado archivo: {excel_path}")

    df = pd.read_excel(excel_path)
    existing = set(df["Numero de Seguimiento/Orden"].astype(str))
    new_shipments: List[Shipment] = []
    cruz_del_sur_track = None

    for file in os.listdir(args.directory):
        path = os.path.join(args.directory, file)
        lower = file.lower()
        if lower.endswith(".pdf"):
            if "fedex" in lower:
                items = extract_fedex_pdf(path)
                print(f"FedEx: {len(items)} envíos de {file}")
                new_shipments.extend(items)
            elif "manifiesto" in lower or "correos" in lower:
                items = extract_correos_chile_pdf(path)
                print(f"CorreosChile: {len(items)} envíos de {file}")
                new_shipments.extend(items)
        elif lower.endswith((".xlsx", ".xls")):
            if "cruz" in lower:
                items = extract_cruz_del_sur_excel(path)
                print(f"Cruz del Sur: {len(items)} envíos de {file}")
                new_shipments.extend(items)
                for it in items:
                    if it.tracking_number:
                        cruz_del_sur_track = it.tracking_number
                        break
            elif "starken" in lower:
                items = extract_starken_excel(path)
                print(f"Starken: {len(items)} envíos de {file}")
                new_shipments.extend(items)

    new_filtered = [s for s in new_shipments if s.tracking_number not in existing]
    if new_filtered:
        df2 = pd.DataFrame([
            [s.carrier, s.tracking_number, s.consignee, s.company, s.reference, s.status]
            for s in new_filtered
        ], columns=df.columns)
        df = pd.concat([df, df2], ignore_index=True)
        print("Actualizando estados...")

    cruz_update = None
    if cruz_del_sur_track:
        estado = consulta_cruz_del_sur(cruz_del_sur_track)
        if estado:
            cruz_update = (cruz_del_sur_track, estado)

    updated_rows = [update_status(Shipment(**row._asdict()), cruz_update) for row in df.itertuples(index=False)]
    df_updated = pd.DataFrame([
        [s.carrier, s.tracking_number, s.consignee, s.company, s.reference, s.status]
        for s in updated_rows
    ], columns=df.columns)
    df_updated.to_excel(excel_path, index=False)
    print("Excel actualizado")


if __name__ == "__main__":  # pragma: no cover - CLI
    main()
