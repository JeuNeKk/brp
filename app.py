import io
import re
import unicodedata

import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

APP_TITLE = "BRP: Excel imprimible por docente"
TEMPLATE_PATH = "plantilla.xlsx"
DATA_ROW = 160
PRINT_AREA = "A162:C193"
MAX_REGISTROS = 300


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


def leer_registros(base_bytes):
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
    rbd_col = col.get("rbd (establecimiento)")
    periodo_col = col.get("período") or col.get("periodo")
    mes_col = col.get("mes")
    anio_col = col.get("año") or col.get("ano")

    if not rut_col:
        raise ValueError("Falta la columna 'RUT (Docente)'.")

    n_cols = ws_base.max_column
    registros = []

    rbd_detectado = ""
    periodo_detectado = ""

    for r in range(2, ws_base.max_row + 1):
        rut = ws_base.cell(r, rut_col).value

        if rut in (None, ""):
            continue

        if not rbd_detectado and rbd_col:
            rbd_detectado = ws_base.cell(r, rbd_col).value or ""

        if not periodo_detectado:
            if periodo_col:
                periodo_detectado = ws_base.cell(r, periodo_col).value or ""
            elif mes_col and anio_col:
                mes = ws_base.cell(r, mes_col).value or ""
                anio = ws_base.cell(r, anio_col).value or ""
                periodo_detectado = f"{mes} {anio}".strip()

        ap1 = ws_base.cell(r, ap1_col).value if ap1_col else ""
        ap2 = ws_base.cell(r, ap2_col).value if ap2_col else ""
        nom = ws_base.cell(r, nom_col).value if nom_col else ""

        registros.append(
            {
                "rut": rut,
                "ap1": ap1,
                "ap2": ap2,
                "nom": nom,
                "row_values": [
                    ws_base.cell(r, c).value for c in range(1, n_cols + 1)
                ],
            }
        )

    if not registros:
        raise ValueError("El archivo no contiene registros con RUT.")

    registros.sort(
        key=lambda x: (
            sort_key(x["ap1"]),
            sort_key(x["ap2"]),
            sort_key(x["nom"]),
            str(x["rut"]).strip(),
        )
    )

    return registros, rbd_detectado, periodo_detectado


def procesar(base_bytes):
    registros, _, _ = leer_registros(base_bytes)

    if len(registros) > MAX_REGISTROS:
        raise ValueError(
            f"Se detectaron {len(registros)} registros. "
            "Este archivo parece ser histórico y no mensual. "
            "Por favor suba el archivo BRP mensual correcto."
        )

    wb_out = openpyxl.load_workbook(TEMPLATE_PATH, data_only=False)
    tmpl = wb_out.active

    wb_out.calculation.fullCalcOnLoad = True
    wb_out.calculation.forceFullCalc = True
    wb_out.calculation.calcMode = "auto"

    progress = st.progress(0)

    for idx, reg in enumerate(registros, start=1):
        ws = wb_out.copy_worksheet(tmpl)

        ws.title = safe_sheet_title(
            f'{reg["ap1"]}_{reg["ap2"]}_{reg["rut"]}',
            set(wb_out.sheetnames),
        )

        for c, val in enumerate(reg["row_values"], start=1):
            ws.cell(DATA_ROW, c).value = val

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
            footer=0.3,
        )

        for rr in range(1, 162):
            ws.row_dimensions[rr].hidden = True

        for cc in range(4, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(cc)].hidden = True

        ws.sheet_view.showGridLines = False
        ws.sheet_view.selection[0].activeCell = "A162"
        ws.sheet_view.selection[0].sqref = "A162"
        ws.sheet_view.topLeftCell = "A162"

        progress.progress(idx / len(registros))

    wb_out.remove(tmpl)

    output = io.BytesIO()
    wb_out.save(output)
    output.seek(0)

    return output.getvalue(), len(registros)


st.set_page_config(page_title=APP_TITLE, page_icon="🧾")

st.title("🧾 " + APP_TITLE)

st.markdown(
    """
**Instrucciones:**

1. Descargue el archivo BRP mensual.
2. Súbalo aquí sin modificarlo.
3. Revise que el período y cantidad de docentes sean correctos.
4. Descargue el Excel final listo para imprimir.

✅ **Orden de impresión:** Apellido paterno A → Z.
"""
)

uploaded = st.file_uploader("Sube el Excel mensual (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        file_bytes = uploaded.getvalue()
        registros_preview, rbd, periodo = leer_registros(file_bytes)
        total_preview = len(registros_preview)

        st.success("Archivo cargado correctamente.")

        col1, col2, col3 = st.columns(3)
        col1.metric("RBD detectado", rbd if rbd else "No detectado")
        col2.metric("Período", periodo if periodo else "No detectado")
        col3.metric("Docentes encontrados", total_preview)

        if total_preview > MAX_REGISTROS:
            st.error(
                f"Se detectaron {total_preview} registros. "
                "Este archivo parece ser histórico y no mensual. "
                "Por seguridad, no se generará el Excel."
            )
        else:
            st.info("Revise los datos detectados. Si son correctos, genere el Excel.")

            if st.button("Generar Excel imprimible"):
                with st.spinner("Generando archivo..."):
                    try:
                        result_bytes, total = procesar(file_bytes)

                        st.success(
                            f"Archivo generado correctamente. Docentes procesados: {total}"
                        )

                        st.download_button(
                            label="⬇️ Descargar BRP_imprimible.xlsx",
                            data=result_bytes,
                            file_name="BRP_imprimible.xlsx",
                            mime=(
                                "application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet"
                            ),
                        )

                    except Exception as e:
                        st.error(f"Error: {e}")

    except Exception as e:
        st.error(f"Error: {e}")
