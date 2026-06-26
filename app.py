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


def num(value):
    if value in (None, ""):
        return 0
    if isinstance(value, (int, float)):
        return value
    try:
        return float(str(value).replace(".", "").replace(",", "."))
    except Exception:
        return 0


def fill_print_values(ws, data):
    def v(name):
        return data.get(normalize(name), "")

    subv_titulo = num(v("Subvención título"))
    trans_titulo = num(v("Transferencia directa título"))

    subv_mencion = num(v("Subvención mención"))
    trans_mencion = num(v("Transferencia directa mención"))

    total_tramo = num(v("Total tramo"))
    asign_prioritarios = num(v("Asignación directa alumnos prioritarios"))

    values_b = {
        162: v("Rbd (Establecimiento)"),
        163: v("RUT (Docente)"),
        164: v("Nombres (Docente)"),
        165: v("Primer Apellido (Docente)"),
        166: v("Segundo Apellido (Docente)"),
        167: v("Bienios"),
        168: v("Tramo"),
        169: v("Carrera docente"),
        170: v("Derecho a pago asignación de tramo"),
        171: v("Derecho a prioritario"),
        172: v("Horas de contrato"),
        173: v("Total días trabajados o descontados"),
        174: v("Subvención título"),
        175: v("Transferencia directa título"),
        176: v("Subvención mención"),
        177: v("Transferencia directa mención"),
        178: v("Total subvención reconocimiento profesional"),
        179: v("Total transferencia directa reconocimiento"),
        180: v("Total reconocimiento profesional"),
        181: v("Subvención tramo"),
        182: v("Transferencia directa tramo"),
        183: v("Total tramo"),
        184: v("Asignación directa alumnos prioritarios"),
        185: v("Total subvenciones"),
        186: v("Total transferencia directa"),
        187: v("Total Asignación por Desem Dificil"),
        188: v("A pagar docente desempeño difícil"),
        189: v("Período"),
        190: v("Tipo de pago"),
        191: v("Mes"),
        192: v("Año"),
        193: v("Porcentaje Alumnos Prioritarios"),
    }

    values_c = {
        162: "VALORES",
        175: subv_titulo + trans_titulo,
        177: subv_mencion + trans_mencion,
        183: total_tramo,
        184: asign_prioritarios,
    }

    for row in range(162, 194):
        ws.cell(row, 2).value = values_b.get(row, "")
        ws.cell(row, 3).value = values_c.get(row, "")


def procesar(base_bytes):
    wb_base = openpyxl.load_workbook(io.BytesIO(base_bytes), data_only=False)
    ws_base, col = find_base_sheet(wb_base)

    if ws_base is None:
        raise ValueError(
            "No se encontró la columna 'RUT (Docente)'. "
            "Suba el archivo BRP mensual correcto."
        )

    rut_col = col.get("rut (docente)")
    ap1_col = col.get("primer apellido (docente)")
    ap2_col = col.get("segundo apellido (docente)")
    nom_col = col.get("nombres (docente)")

    registros = []

    for r in range(2, ws_base.max_row + 1):
        rut = ws_base.cell(r, rut_col).value

        if rut in (None, ""):
            continue

        data = {}
        for header, c in col.items():
            data[header] = ws_base.cell(r, c).value

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

    for idx, reg in enumerate(registros, start=1):
        ws = wb_out.create_sheet(
            safe_sheet_title(
                f'{reg["ap1"]}_{reg["ap2"]}_{reg["rut"]}',
                set(wb_out.sheetnames)
            )
        )

        for row in range(162, 194):
            for col_num in range(1, 4):
                copy_cell(
                    ws_template.cell(row, col_num),
                    ws.cell(row, col_num)
                )

        for col_num in range(1, 4):
            letter = get_column_letter(col_num)
            ws.column_dimensions[letter].width = (
                ws_template.column_dimensions[letter].width
            )

        for row in range(162, 194):
            ws.row_dimensions[row].height = (
                ws_template.row_dimensions[row].height
            )

        fill_print_values(ws, reg["data"])

        for rr in range(1, 162):
            ws.row_dimensions[rr].hidden = True

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
