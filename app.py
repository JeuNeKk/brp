import io
import re
import unicodedata
import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

APP_TITLE = "BRP: Excel imprimible por docente"
TEMPLATE_PATH = "plantilla.xlsx"   # Debe estar en el repo
DATA_ROW = 160
PRINT_AREA = "A162:C193"

# ------------------ Helpers ------------------
def normalize(text):
    return str(text).strip().lower() if text is not None else ""

def strip_accents(s: str) -> str:
    """Convierte 'Álvarez' -> 'Alvarez' para ordenar alfabéticamente de forma natural."""
    s = "" if s is None else str(s)
    s = s.strip()
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", s)
        if not unicodedata.combining(ch)
    )

def sort_key_spanish(s: str) -> str:
    # minúsculas + sin acentos para orden
    return strip_accents(s).lower()

def safe_sheet_title(name, existing):
    # Excel no permite : \ / ? * [ ] y máx 31 caracteres
    name = re.sub(r'[:\\/?*\[\]]', ' ', str(name)).strip()
    if not name:
        name = "Hoja"
    name = name[:31]
    base = name
    i = 1
    while name in existing:
        suffix = f"_{i}"
        name = (base[:31 - len(suffix)] + suffix)
        i += 1
    return name

def find_base_sheet(workbook):
    """Encuentra la hoja que tenga la columna 'RUT (Docente)' en la fila 1."""
    for ws in workbook.worksheets:
        headers = [normalize(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]
        if "rut (docente)" in headers:
            return ws, {headers[i]: i + 1 for i in range(len(headers)) if headers[i]}
    return None, None

# ------------------ Core ------------------
def procesar(base_bytes: bytes) -> bytes:
    # 1) Leer base mensual (subida por usuario)
    wb_base = openpyxl.load_workbook(io.BytesIO(base_bytes), data_only=False)

    ws_base, col = find_base_sheet(wb_base)
    if ws_base is None:
        raise ValueError("No se encontró la columna 'RUT (Docente)'. Suba el archivo BRP mensual correcto.")

    rut_col = col.get("rut (docente)")
    nom_col = col.get("nombres (docente)")
    ap1_col = col.get("primer apellido (docente)")
    ap2_col = col.get("segundo apellido (docente)")

    if not rut_col:
        raise ValueError("Falta la columna 'RUT (Docente)' en la fila 1.")

    n_cols = ws_base.max_column

    # Validación mínima: al menos 1 RUT desde fila 2
    has_any = False
    for r in range(2, ws_base.max_row + 1):
        if ws_base.cell(r, rut_col).value not in (None, ""):
            has_any = True
            break
    if not has_any:
        raise ValueError("El archivo no contiene registros (no se encontraron RUT).")

    # 2) Cargar plantilla del repo/servidor
    wb_out = openpyxl.load_workbook(TEMPLATE_PATH, data_only=False)
    tmpl = wb_out.active

    # Forzar recálculo al abrir en Excel
    wb_out.calculation.fullCalcOnLoad = True
    wb_out.calculation.calcMode = "auto"

    # 3) Recolectar registros
    registros = []
    for r in range(2, ws_base.max_row + 1):
        rut = ws_base.cell(r, rut_col).value
        if rut in (None, ""):
            continue

        ap1 = ws_base.cell(r, ap1_col).value if ap1_col else ""
        ap2 = ws_base.cell(r, ap2_col).value if ap2_col else ""
        nom = ws_base.cell(r, nom_col).value if nom_col else ""

        registros.append({
            "rut": rut,
            "ap1": ap1,
            "ap2": ap2,
            "nom": nom,
            "row_values": [ws_base.cell(r, c).value for c in range(1, n_cols + 1)]
        })

    # 4) ORDENAR por Apellido Paterno (A→Z), sin acentos; tie-breakers: ap2, nombres, rut
    registros.sort(
        key=lambda x: (
            sort_key_spanish(x["ap1"]),
            sort_key_spanish(x["ap2"]),
            sort_key_spanish(x["nom"]),
            str(x["rut"]).strip()
        )
    )

    # 5) Crear hojas en ese orden
    for reg in registros:
        ws = wb_out.copy_worksheet(tmpl)

        sheet_name = f'{reg["rut"]}_{reg["ap1"]}_{reg["ap2"]}_{reg["nom"]}'
        ws.title = safe_sheet_title(sheet_name, set(wb_out.sheetnames))

        # pegar fila completa en la fila buffer 160
        for c, val in enumerate(reg["row_values"], start=1):
            ws.cell(DATA_ROW, c).value = val

        # Configuración impresión
        ws.print_area = PRINT_AREA
        ws.page_setup.orientation = "portrait"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.page_setup.paperSize = 9  # A4
        ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.5, bottom=0.5, header=0.3, footer=0.3)

        # Ocultar lo irrelevante (pero se mantiene para fórmulas)
        for rr in range(1, 162):
            ws.row_dimensions[rr].hidden = True
        for cc in range(4, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(cc)].hidden = True

        # Vista limpia
        ws.sheet_view.showGridLines = False
        ws.sheet_view.selection[0].activeCell = "A162"
        ws.sheet_view.selection[0].sqref = "A162"
        ws.sheet_view.topLeftCell = "A162"

    # Eliminar plantilla original
    wb_out.remove(tmpl)

    # Guardar resultado en memoria
    output = io.BytesIO()
    wb_out.save(output)
    return output.getvalue()

# ------------------ UI ------------------
st.set_page_config(page_title=APP_TITLE, page_icon="🧾")
st.title("🧾 " + APP_TITLE)

st.markdown("""
**Instrucciones (3 pasos):**
1. Descarga el archivo BRP mensual (*Web - Sostenedor - Lista de asignaciones…*).
2. Súbelo aquí **sin modificarlo**.
3. Descarga el Excel final: viene con **1 hoja por docente**, listo para imprimir.

✅ **Orden de impresión:** Apellido paterno (A → Z).
""")

uploaded = st.file_uploader("Sube el Excel mensual (.xlsx)", type=["xlsx"])

if uploaded:
    st.success("Archivo cargado.")
    if st.button("Generar Excel imprimible"):
        try:
            result_bytes = procesar(uploaded.getvalue())
            st.download_button(
                label="⬇️ Descargar BRP_imprimible.xlsx",
                data=result_bytes,
                file_name="BRP_imprimible.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error: {e}")