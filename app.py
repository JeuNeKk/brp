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
DATA_ROW = 160
PRINT_AREA = "A162:C193"


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


def copy_cell(src, dst):
    dst.value = src.value

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

        if "rut (docente)" in headers:
            return ws, {
                headers[i]: i + 1
                for i in range(len(headers))
                if headers[i]
            }

    return None, None


def procesar(base_bytes):
    wb_base = openpyxl.load_workbook(io.BytesIO(base_bytes), data_only=False)
    ws_base, col = find_base_sheet(wb_base)

    if ws_base is None:
        raise ValueError(
            "No se encontró la columna 'RUT (Docente)'. "
            "Suba el archivo BRP mensual correcto."
        )

    rut_col = col.get("rut (docente)")
    nom_col = col.get("nombres (docente)")
    ap1_col = col.get("primer apellido (docente)")
    ap2_col = col.get("segundo apellido (docente)")

    n_cols = ws_base.max_column
    registros = []

    for r in range(2, ws_base.max_row + 1):
        rut = ws_base.cell(r, rut_col).value

        if rut in (None, ""):
            continue

        registros.append({
            "rut": rut,
            "ap1": ws_base.cell(r, ap1_col).value if ap1_col else "",
            "ap2": ws_base.cell(r, ap2_col).value if ap2_col else "",
            "nom": ws_base.cell(r, nom_col).value if nom_col else "",
            "row_values": [
                ws_base.cell(r, c).value
                for c in range(1, n_cols + 1)
            ]
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

    wb_out.calculation.fullCalcOnLoad = True
    wb_out.calculation.calcMode = "auto"

    progress = st.progress(0)

    for idx, reg in enumerate(registros, start=1):
        ws = wb_out.create_sheet(
            safe_sheet_title(
                f'{reg["ap1"]}_{reg["ap2"]}_{reg["rut"]}',
                set(wb_out.sheetnames)
            )
        )

        # Copiar solo el bloque necesario: A162:C193
        for row in range(162, 194):
            for col_num in range(1, 4):
                copy_cell(
                    ws_template.cell(row, col_num),
                    ws.cell(row, col_num)
                )

        # Copiar anchos A:C
        for col_num in range(1, 4):
            letter = get_column_letter(col_num)
            ws.column_dimensions[letter].width = (
                ws_template.column_dimensions[letter].width
            )

        # Copiar alturas 162:193
        for row in range(162, 194):
            ws.row_dimensions[row].height = (
                ws_template.row_dimensions[row].height
            )

        # Pegar datos en fila buffer 160
        for c, val in enumerate(reg["row_values"], start=1):
            ws.cell(DATA_ROW, c).value = val

        # Ocultar fila buffer y filas superiores
        for rr in range(1, 162):
            ws.row_dimensions[rr].hidden = True

        # Ocultar columnas desde D en adelante
        for cc in range(4, max(n_cols, 40) + 1):
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

                st.success(f"Archivo generado correctamente. Docentes procesados: {total}")

                st.download_button(
                    label="⬇️ Descargar BRP_imprimible.xlsx",
                    data=result_bytes,
                    file_name="BRP_imprimible.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Error: {e}")
