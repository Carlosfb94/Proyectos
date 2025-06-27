import os, io, json, base64, zipfile, textwrap
from datetime import datetime

# ---------- 1. Pega aquí el FLOW JSON completo ----------
FLOW_JSON = """
{ "$schema":"https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2025-06-01/workflowdefinition.json#",
  "contentVersion":"1.0.0.0",
  "parameters":{},
  "triggers":{  /* ... TU BLOQUE COMPLETO ... */ },
  "actions":{   /* ...                           */ },
  "outputs":{},
  "connectionReferences":{ /* ... */ }
}
""".strip()
# ---------- 2. Variables de solución ----------
SOLUTION_NAME   = "ActualizarStock_Compras_Solution"
DISPLAY_NAME    = "Actualizar Stock – Compras"
FLOW_GUID       = "{11111111-2222-3333-4444-555555555555}"
PUBLISHER_UNAME = "codx"
PUBLISHER_PREF  = "codx"
VERSION         = "1.0.0.0"

# ---------- 3. Contenidos de los XML ----------
content_types = textwrap.dedent("""\
    <?xml version="1.0" encoding="utf-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="xml"  ContentType="application/xml"/>
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Override PartName="/Workflows/ActualizarStock_Compras/definition.json" ContentType="application/json"/>
    </Types>
""")

solution_xml = textwrap.dedent(f"""\
    <?xml version="1.0" encoding="utf-8"?>
    <ImportExportXml>
      <SolutionManifest>
        <UniqueName>{SOLUTION_NAME}</UniqueName>
        <Version>{VERSION}</Version>
        <Publisher>
          <UniqueName>{PUBLISHER_UNAME}</UniqueName>
          <Prefix>{PUBLISHER_PREF}</Prefix>
        </Publisher>
        <Managed>0</Managed>
        <RootComponents>
          <RootComponent type="29" id="{FLOW_GUID}" />
        </RootComponents>
      </SolutionManifest>
      <Workflows>
        <Workflow path="Workflows/ActualizarStock_Compras/definition.json" id="{FLOW_GUID}" />
      </Workflows>
    </ImportExportXml>
""")

customizations_xml = textwrap.dedent(f"""\
    <?xml version="1.0" encoding="utf-8"?>
    <ImportExportXml>
      <Workflows>
        <Workflow>
          <WorkflowId>{FLOW_GUID}</WorkflowId>
          <Name>ActualizarStock_Compras</Name>
          <DisplayName>{DISPLAY_NAME}</DisplayName>
          <Category>5</Category>
        </Workflow>
      </Workflows>
    </ImportExportXml>
""")

# ---------- 4. Crear ZIP en memoria ----------
buffer = io.BytesIO()
with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("solution.xml",         solution_xml)
    z.writestr("Other/Customizations.xml", customizations_xml)
    z.writestr("Workflows/ActualizarStock_Compras/definition.json", json.dumps(json.loads(FLOW_JSON), indent=2))
zip_bytes = buffer.getvalue()

# ---------- 5. Guardar ZIP en disco ----------
zip_filename = "ActualizarStock.zip"
with open(zip_filename, "wb") as f:
    f.write(zip_bytes)
print(f"✅ ZIP escrito: {os.path.abspath(zip_filename)}  ({len(zip_bytes)//1024} KB)")

# ---------- 6. Imprimir Base-64 (opcional) ----------
b64 = base64.b64encode(zip_bytes).decode()
print("\n--- ZIP BASE64 ---")
print(b64)
print("\n--- FIN ---")
