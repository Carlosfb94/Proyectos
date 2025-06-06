import os
import pdfplumber
import pandas as pd
import re
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from PIL import Image

CARPETA_DIA = r'C:\Users\cfernandez\OneDrive - Llacolenvet\Escritorio\Python\25 Mayo'
ARCHIVO_EXCEL_DIA = os.path.join(CARPETA_DIA, '25 Mayo_envios.xlsx')
API_KEY_2CAPTCHA = 'c109e2e4c8813bc91d4253d5e7cd8a5f'
COORDS = (563, 409, 701, 471)

COLUMNAS = [
    'Tipo',
    'Numero de Seguimiento/Orden',
    'Consignatario/Destinatario',
    'Compañía de Envío',
    'Referencia',
    'Estado'
]

def buscar_tabla_excel_generica(excel_path, headers_objetivo):
    df = pd.read_excel(excel_path, header=None)
    headers_objetivo = [h.upper() for h in headers_objetivo]
    for idx, row in df.iterrows():
        row_vals = [str(x).strip().upper() for x in row]
        if all(encabezado in row_vals for encabezado in headers_objetivo):
            df_tabla = df.iloc[idx+1:].copy()
            df_tabla.columns = row_vals
            df_tabla = df_tabla.reset_index(drop=True)
            return df_tabla
    return None

def extraer_starken_excel(excel_path):
    for headers in [
        ['ORDEN DE TRANSPORTE', 'DESTINATARIO'],
        ['ORDEN TRANSPORTE', 'DESTINATARIO'],
        ['NUMERO DE SEGUIMIENTO', 'DESTINATARIO']
    ]:
        df_tabla = buscar_tabla_excel_generica(excel_path, headers)
        if df_tabla is not None:
            col_orden = [c for c in df_tabla.columns if "ORDEN" in c or "NUMERO" in c][0]
            col_dest = [c for c in df_tabla.columns if "DESTINATARIO" in c][0]
            return [
                [
                    'Starken',
                    str(fila.get(col_orden, '')).strip(),
                    str(fila.get(col_dest, '')).strip(),
                    'Starken',
                    '',
                    ''
                ]
                for _, fila in df_tabla.iterrows()
                if str(fila.get(col_orden, '')).strip().lower() not in ['', 'nan']
            ]
    print(f"❌ No se encontraron encabezados válidos en: {excel_path}")
    return []

def extraer_fedex(pdf_path):
    envios = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                for fila in table[1:]:
                    if len(fila) >= 3 and fila[0] and fila[2]:
                        nro_seguimiento = fila[0].strip()
                        consignatario = fila[2].strip().split("\n")[0]
                        if nro_seguimiento.isdigit() and len(nro_seguimiento) >= 8:
                            envios.append([
                                'FedEx',
                                nro_seguimiento,
                                consignatario,
                                'FedEx',
                                '',
                                ''
                            ])
    return envios

def referencia_valida(ref_string):
    codigos = re.findall(r'F-(\d+)', ref_string)
    for cod in codigos:
        if cod.startswith("36"):
            return True
    return False

def extraer_correos_chile(pdf_path):
    envios = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                headers = [col.strip().upper() for col in table[0]]
                try:
                    idx_destinatario = headers.index('DESTINATARIO')
                    idx_referencia = headers.index('REFERENCIA')
                    idx_seguimiento = headers.index('SEGUIMIENTO')
                except ValueError:
                    continue
                for fila in table[1:]:
                    if (len(fila) > max(idx_destinatario, idx_referencia, idx_seguimiento)):
                        destinatario = fila[idx_destinatario].strip()
                        referencia = fila[idx_referencia].strip()
                        nro_seguimiento = fila[idx_seguimiento].strip()
                        if referencia_valida(referencia) and nro_seguimiento.isdigit():
                            envios.append([
                                'Correos de Chile',
                                nro_seguimiento,
                                destinatario,
                                'Correos de Chile',
                                referencia,
                                ''
                            ])
    return envios

def extraer_cruz_del_sur_excel(excel_path):
    df_tabla = buscar_tabla_excel_generica(excel_path, ['ORDEN TRANSPORTE', 'DESTINATARIO'])
    if df_tabla is None:
        print(f"❌ No se encontraron encabezados válidos en: {excel_path}")
        return []
    return [
        [
            'Cruz del Sur',
            str(fila.get('ORDEN TRANSPORTE', '')).strip(),
            str(fila.get('DESTINATARIO', '')).strip(),
            'Cruz del Sur',
            '',
            ''
        ]
        for _, fila in df_tabla.iterrows()
        if str(fila.get('ORDEN TRANSPORTE', '')).strip().lower() not in ['', 'nan']
    ]

def obtener_estado_fedex(nro):
    url = f"https://clsclweb.tntchile.cl/txapgw/tracking.asp?boleto={nro}"
    try:
        r = requests.get(url, timeout=15)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, 'html.parser')
        situacion = soup.find(string=lambda text: text and ("Situacion:" in text or "Situación:" in text))
        if situacion:
            return situacion.split(":")[-1].strip()
        elif "ENTREGADA" in r.text.upper():
            return "Entregada"
        else:
            return "No disponible"
    except Exception as e:
        return f"Error: {str(e)}"

def obtener_estado_correos(nro):
    url = f"https://www.correos.cl/web/guest/seguimiento-en-linea?numero={nro}"
    try:
        r = requests.get(url, timeout=20)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, 'html.parser')
        estado_elem = soup.find('span', {'class': 'estado'})
        if estado_elem and estado_elem.text.strip():
            return estado_elem.text.strip()
        if "ENTREGADO" in r.text.upper():
            return "Entregado"
        elif "NO REGISTRA INFORMACIÓN" in r.text.upper():
            return "No registra información"
        else:
            return "En tránsito o no disponible"
    except Exception as e:
        return f"Error: {str(e)}"

def obtener_estado_starken(nro):
    url = f"https://www.starken.cl/seguimiento?codigo={nro}"
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--window-size=1920,1080")
    driver = None
    try:
        driver = webdriver.Chrome(options=chrome_options)
        driver.get(url)
        wait = WebDriverWait(driver, 20)
        main_text = ""
        date_text = ""

        try:
            # Busca el estado de entrega principal
            estado_elem = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(text(),'El envío ya fue entregado')]")
                )
            )
            main_text = estado_elem.text.strip()
        except Exception:
            # Si no fue entregado, intenta con otros posibles estados
            posibles = [
                "En tránsito", "En reparto", "En sucursal destino",
                "Recibido por Starken", "Solicitud de Envío Creado"
            ]
            for estado in posibles:
                try:
                    elem = driver.find_element(By.XPATH, f"//*[contains(text(),'{estado}')]")
                    if elem and elem.text.strip():
                        main_text = elem.text.strip()
                        break
                except Exception:
                    continue

        # Busca la fecha de entrega
        try:
            fecha_elem = driver.find_element(By.XPATH, "//*[contains(text(),'Entregado con fecha')]")
            date_text = fecha_elem.text.strip()
        except Exception:
            date_text = ""

        # Extrae la fecha limpia
        match = re.search(r'(\d{2}-\d{2}-\d{4} \d{2}:\d{2}:\d{2})', date_text)
        if not match:
            match = re.search(r'(\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2})', date_text)
        solo_fecha = match.group(1) if match else ""

        if main_text and solo_fecha:
            return f"{main_text} - {solo_fecha}"
        elif main_text:
            return main_text
        else:
            print("DEBUG HTML STARKEN:", driver.page_source[:2000])
            return "No se detectó estado claro."
    except Exception as e:
        return f"Error Selenium: {str(e)}"
    finally:
        if driver:
            driver.quit()

def extraer_fecha_estado(tabla_text):
    resultados = []
    for linea in tabla_text.split('\n'):
        match = re.match(r'(\d{2}/\d{2}/\d{4} \d{2}:\d{2})\s+(.*)', linea)
        if match:
            fecha_str = match.group(1)
            resto = match.group(2).strip()
            try:
                fecha_dt = datetime.strptime(fecha_str, "%d/%m/%Y %H:%M")
                resultados.append((fecha_dt, linea.strip(), resto))
            except Exception:
                pass
    return resultados

def consulta_cruz_del_sur(nro_doc, max_intentos=5):
    print(f"Consultando Cruz del Sur para {nro_doc}...")
    intentos = 0
    while intentos < max_intentos:
        try:
            driver = webdriver.Chrome()
            driver.get("https://www.cruzdelsurcarga.cl/seguimiento/")
            time.sleep(2)
            input_nrodoc = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "nrodoc"))
            )
            input_nrodoc.clear()
            input_nrodoc.send_keys(nro_doc)
            driver.save_screenshot('screenshot.png')
            im = Image.open('screenshot.png')
            captcha_im = im.crop(COORDS)
            captcha_im.save('captcha_crop.png')
            print(f"Enviando captcha a 2Captcha... (Intento {intentos+1})")
            with open('captcha_crop.png', 'rb') as captcha_file:
                files = {'file': captcha_file}
                data = {'key': API_KEY_2CAPTCHA, 'method': 'post'}
                try:
                    r = requests.post('http://2captcha.com/in.php', files=files, data=data, timeout=20)
                    result = r.text
                except Exception as e:
                    print(f"Error conexión 2Captcha: {e}")
                    driver.quit()
                    intentos += 1
                    continue
            if 'OK|' not in result:
                print('Error enviando captcha:', result)
                driver.quit()
                intentos += 1
                continue
            captcha_id = result.split('|')[1]
            print("Captcha enviado correctamente. ID:", captcha_id)
            print("Consultando resultado...")
            captcha_result = None
            for _ in range(15):
                time.sleep(5)
                try:
                    res = requests.get(f"http://2captcha.com/res.php?key={API_KEY_2CAPTCHA}&action=get&id={captcha_id}", timeout=20)
                except Exception as e:
                    print(f"Error conexión 2Captcha: {e}")
                    driver.quit()
                    intentos += 1
                    continue
                if res.text == "CAPCHA_NOT_READY":
                    print("No listo aún, esperando 5 segundos...")
                    continue
                elif "OK|" in res.text:
                    captcha_result = res.text.split('|')[1]
                    break
                else:
                    print("Error obteniendo resultado:", res.text)
                    driver.quit()
                    intentos += 1
                    continue
            if not captcha_result:
                print("No se pudo resolver el captcha.")
                driver.quit()
                intentos += 1
                continue
            print("¡Captcha resuelto:", captcha_result)
            input_captcha = driver.find_element(By.ID, "captcha")
            input_captcha.clear()
            input_captcha.send_keys(captcha_result)
            consultar_btn = None
            inputs = driver.find_elements(By.TAG_NAME, 'input')
            for b in inputs:
                value = b.get_attribute('value') or ''
                if 'CONSULTAR' in value.upper():
                    consultar_btn = b
                    break
            if not consultar_btn:
                buttons = driver.find_elements(By.TAG_NAME, 'button')
                for b in buttons:
                    if 'CONSULTAR' in (b.text or '').upper():
                        consultar_btn = b
                        break
            if not consultar_btn:
                print("No se encontró ningún botón 'CONSULTAR'.")
                driver.quit()
                intentos += 1
                continue
            consultar_btn.click()
            print("¡Consulta enviada! Esperando el resultado...")
            time.sleep(6)
            estado_final = None
            try:
                tablas = driver.find_elements(By.TAG_NAME, "table")
                todas_las_fechas = []
                for idx, tabla in enumerate(tablas):
                    tabla_text = tabla.text
                    resultados = extraer_fecha_estado(tabla_text)
                    todas_las_fechas.extend(resultados)
                if todas_las_fechas:
                    todas_las_fechas.sort(key=lambda x: x[0], reverse=True)
                    fecha_mas_reciente, linea, estado = todas_las_fechas[0]
                    print(">>> Resultado más reciente Cruz del Sur:")
                    print(f"Fecha: {fecha_mas_reciente} | Estado: {estado}")
                    estado_final = f"{estado} [{fecha_mas_reciente.strftime('%d/%m/%Y %H:%M')}]"
                    driver.quit()
                    return estado_final
                else:
                    print("No se encontraron estados con fechas reconocidas.")
            except Exception as e:
                print("Error buscando tablas:", e)
            driver.quit()
            intentos += 1
        except Exception as e:
            print(f"Fallo en la consulta Cruz del Sur: {e}")
            intentos += 1
        print("Refrescando e intentando de nuevo...\n")
        time.sleep(3)
    print(f"Falló la consulta Cruz del Sur después de {max_intentos} intentos.")
    return None

def actualizar_estados(df, cruz_del_sur_update):
    for i, row in df.iterrows():
        tipo = str(row['Tipo']).lower()
        nro = str(row['Numero de Seguimiento/Orden'])
        if tipo == "fedex":
            estado = obtener_estado_fedex(nro)
        elif tipo == "correos de chile":
            estado = obtener_estado_correos(nro)
        elif tipo == "starken":
            estado = obtener_estado_starken(nro)
        elif tipo == "cruz del sur":
            if cruz_del_sur_update and nro == cruz_del_sur_update[0]:
                estado = cruz_del_sur_update[1]
            else:
                estado = row['Estado'] if 'Estado' in row else 'Requiere consulta manual'
        else:
            estado = "Sin definir"
        df.at[i, 'Estado'] = estado
        print(f"{tipo.title()} {nro}: {estado}")
    return df

def main():
    if not os.path.exists(ARCHIVO_EXCEL_DIA):
        pd.DataFrame(columns=COLUMNAS).to_excel(ARCHIVO_EXCEL_DIA, index=False)
        print(f"Creado archivo: {ARCHIVO_EXCEL_DIA}")

    df = pd.read_excel(ARCHIVO_EXCEL_DIA)
    nuevos = []
    cruz_del_sur_num = None

    for archivo in os.listdir(CARPETA_DIA):
        ruta = os.path.join(CARPETA_DIA, archivo)
        if archivo.lower().endswith('.pdf'):
            if 'fedex' in archivo.lower():
                extraidos = extraer_fedex(ruta)
                print(f'FedEx: {len(extraidos)} envíos extraídos de {archivo}')
                nuevos += extraidos
            elif 'manifiesto' in archivo.lower() or 'correos' in archivo.lower():
                extraidos = extraer_correos_chile(ruta)
                print(f'CorreosChile: {len(extraidos)} envíos extraídos de {archivo}')
                nuevos += extraidos
        elif archivo.lower().endswith(('.xlsx', '.xls')):
            if 'cruz' in archivo.lower():
                extraidos = extraer_cruz_del_sur_excel(ruta)
                print(f'Cruz del Sur: {len(extraidos)} envíos extraídos de {archivo}')
                nuevos += extraidos
                for envio in extraidos:
                    if envio[1]:
                        cruz_del_sur_num = envio[1]
                        break
            elif 'starken' in archivo.lower():
                extraidos = extraer_starken_excel(ruta)
                print(f'Starken: {len(extraidos)} envíos extraídos de {archivo}')
                nuevos += extraidos

    existentes = set(df['Numero de Seguimiento/Orden'].astype(str))
    nuevos_final = [n for n in nuevos if n[1] not in existentes and n[1] not in ['', 'nan']]
    if nuevos_final:
        for fila in nuevos_final:
            print(f"Agregado: {fila}")
        df2 = pd.DataFrame(nuevos_final, columns=COLUMNAS)
        df = pd.concat([df, df2], ignore_index=True)
        print("Actualizando estados, esto puede demorar...")

    cruz_del_sur_estado = None
    if cruz_del_sur_num:
        cruz_del_sur_estado = consulta_cruz_del_sur(cruz_del_sur_num, max_intentos=5)
        if cruz_del_sur_estado:
            cruz_del_sur_update = (cruz_del_sur_num, cruz_del_sur_estado)
        else:
            cruz_del_sur_update = None
    else:
        cruz_del_sur_update = None

    df = actualizar_estados(df, cruz_del_sur_update)
    df.to_excel(ARCHIVO_EXCEL_DIA, index=False)
    print("Excel del día actualizado.")

if __name__ == "__main__":
    main()