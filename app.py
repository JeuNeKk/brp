import io
import re
import unicodedata
from copy import copy

import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

APP_TITLE = "BRP: Excel imprimible por docente"
TEMPLATE_PATH = "plantilla.xlsx"
PRINT_AREA = "A162:C193"

COLUMNAS_BRPF = [
    "Rbd (Establecimiento)",
    "RUT (Docente)",
    "Nombres (Docente)",
    "Primer Apellido (Docente)",
    "Segundo Apellido (Docente)",
    "Bienios",
    "Tramo",
    "Carrera docente",
    "Derecho a pago asignación de tramo",
    "Derecho a prioritario",
    "Horas de contrato",
    "Total días trabajados o descontados",
    "Subvención título",
    "Transferencia directa título",
    "Subvención mención",
    "Transferencia directa mención",
    "Total subvención reconocimiento profesional",
    "Total transferencia directa reconocimiento",
    "Total reconocimiento profesional",
    "Subvención tramo",
    "Transferencia directa tramo",
    "Total tramo",
    "Asignación directa alumnos prioritarios",
    "Total subvenciones",
    "Total transferencia directa",
    "Total Asignación por Desem Dificil",
    "A pagar docente desempeño difícil",
    "Período",
    "Tipo de pago",
    "Mes",
    "Año",
    "Porcentaje Alumnos Prioritarios",
]


def normalize(text):
    return str(text).strip().lower() if text is not None else ""


def strip_accents(text):
    text = "" if text is None else str(text).strip()
    return "".join(
        c for c in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(c)
    )


def sort_key(text):
    return strip_accents(text).lower()


def clean_sheet_name(text):
    text = "" if text is None else str(text).strip()
    text = re.sub(r'[:\\/?*\[\]]', " ", text)
    text = re.sub(r"\s+", " ", text)
    return text


def safe_sheet_title(name, existing):
    name = clean_sheet_name(name) or "Hoja"
    base = name[:31]
    title = base
    i = 1

    while title in existing:
        suffix = f"_{i}"
        title = base[:31 - len(suffix)] + suffix
        i += 1

    return title[:31]


def num(value):
    if value in (None, ""):
        return 0

    if isinstance(value, (int, float)):
        return value

    text = str(value).strip()

    try:
        return float(text)
    except Exception:
        pass

    try:
        return float(text.replace(".", "").replace(",", "."))
    except Exception:
        return 0


def copy_style_only(src, dst):
    dst.value = None

    if src.has_style:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection = copy(src.protection)

    if src.comment:
        dst.comment = copy(src.comment)


def find_base_sheet(workbook):
    for ws in workbook.worksheets:
        headers = [
            normalize(ws.cell(1, c).value)
            for c in range(1, ws.max_column + 1)
        ]

        if normalize("RUT (Docente)") in headers:
            return ws, {
                headers[i]: i + 1
                for i in range(len(headers))
                if headers[i]
            }

    return None, None


def get_value(data, field_name):
    return data.get(normalize(field_name), "")


def crear_hoja_docente(wb_out, ws_template, registro):
    existing = set(wb_out.sheetnames)

    ws = wb_out.create_sheet(
        safe_sheet_title(
            f'{registro["ap1"]}_{registro["ap2"]}_{registro["rut"]}',
            existing
        )
    )

    # Copiar solo estilos del bloque imprimible A162:C193.
    for row in range(162, 194):
        for col in range(1, 4):
            copy_style_only(
                ws_template.cell(row, col),
                ws.cell(row, col)
            )

    # Copiar anchos de columnas A:C.
    for col in range(1, 4):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = (
            ws_template.column_dimensions[letter].width
        )

    # Copiar alturas de filas 162:193.
    for row in range(162, 194):
        ws.row_dimensions[row].height = (
            ws_template.row_dimensions[row].height
        )

    data = registro["data"]

    # Columna A: etiquetas.
    for idx, label in enumerate(COLUMNAS_BRPF, start=162):
        ws.cell(idx, 1).value = label

    # Columna B: valores reales del docente.
    for idx, label in enumerate(COLUMNAS_BRPF, start=162):
        ws.cell(idx, 2).value = get_value(data, label)

    # Columna C: campos calculados/visuales.
    subv_titulo = num(get_value(data, "Subvención título"))
    trans_titulo = num(get_value(data, "Transferencia directa título"))

    subv_mencion = num(get_value(data, "Subvención mención"))
    trans_mencion = num(get_value(data, "Transferencia directa mención"))

    total_tramo = num(get_value(data, "Total tramo"))
    prioritarios = num(get_value(data, "Asignación directa alumnos prioritarios"))

    ws.cell(162, 3).value = "VALORES"
    ws.cell(175, 3).value = subv_titulo + trans_titulo
    ws.cell(177, 3).value = subv_mencion + trans_mencion
    ws.cell(183, 3).value = total_tramo
    ws.cell(184, 3).value = prioritarios

    # Ocultar todo lo anterior al bloque imprimible.
    for rr in range(1, 162):
        ws.row_dimensions[rr].hidden = True

    # Ocultar columnas desde D.
    for cc in range(4, 40):
        ws.column_dimensions[get_column_letter(cc)].hidden = True

    ws.print_area = PRINT_AREA
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.paperSize = 9

    ws.page_margins = PageMargins(
        left=0.25,
        right=0.25,
        top=0.5,
        bottom=0.5,
        header=0.3,
        footer=0.3
    )

    ws.sheet_view.showGridLines = False
    ws.sheet_view.selection[0].activeCell = "A162"
    ws.sheet_view.selection[0].sqref = "A162"
    ws.sheet_view.topLeftCell = "A162"

    return ws


def procesar(base_bytes):
    wb_base = openpyxl.load_workbook(io.BytesIO(base_bytes), data_only=False)
    ws_base, col = find_base_sheet(wb_base)

    if ws_base is None:
        raise ValueError(
            "No se encontró la columna 'RUT (Docente)'. "
            "Suba el archivo BRP mensual correcto."
        )

    rut_col = col.get(normalize("RUT (Docente)"))
    ap1_col = col.get(normalize("Primer Apellido (Docente)"))
    ap2_col = col.get(normalize("Segundo Apellido (Docente)"))
    nom_col = col.get(normalize("Nombres (Docente)"))

    if not rut_col:
        raise ValueError("Falta la columna 'RUT (Docente)'.")

    registros = []

    for r in range(2, ws_base.max_row + 1):
        rut = ws_base.cell(r, rut_col).value

        if rut in (None, ""):
            continue

        data = {}

        for header_norm, c in col.items():
            data[header_norm] = ws_base.cell(r, c).value

        registros.append({
            "rut": rut,
            "ap1": ws_base.cell(r, ap1_col).value if ap1_col else "",
            "ap2": ws_base.cell(r, ap2_col).value if ap2_col else "",
            "nom": ws_base.cell(r, nom_col).value if nom_col else "",
            "data": data
        })

    if not registros:
        raise ValueError("El archivo no contiene registros con RUT.")

    registros.sort(
        key=lambda x: (
            sort_key(x["ap1"]),
            sort_key(x["ap2"]),
            sort_key(x["nom"]),
            str(x["rut"]).strip()
        )
    )

    wb_template = openpyxl.load_workbook(TEMPLATE_PATH, data_only=False)
    ws_template = wb_template.active

    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)

    progress = st.progress(0)

    for idx, registro in enumerate(registros, start=1):
        crear_hoja_docente(wb_out, ws_template, registro)
        progress.progress(idx / len(registros))

    output = io.BytesIO()
    wb_out.save(output)
    output.seek(0)

    return output.getvalue(), len(registros)


st.set_page_config(page_title=APP_TITLE, page_icon="🧾")
st.title("🧾 " + APP_TITLE)

st.markdown("""
**Instrucciones:**

1. Descargue el archivo BRP mensual.
2. Súbalo aquí sin modificarlo.
3. Descargue el Excel final listo para imprimir.

✅ **Orden de impresión:** Apellido paterno A → Z.
""")

uploaded = st.file_uploader("Sube el Excel mensual (.xlsx)", type=["xlsx"])

if uploaded:
    st.success("Archivo cargado correctamente.")

    if st.button("Generar Excel imprimible"):
        with st.spinner("Generando archivo..."):
            try:
                result_bytes, total = procesar(uploaded.getvalue())

                st.success(
                    f"Archivo generado correctamente. Docentes procesados: {total}"
                )

                st.download_button(
                    label="⬇️ Descargar BRP_imprimible.xlsx",
                    data=result_bytes,
                    file_name="BRP_imprimible.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Error: {e}")
