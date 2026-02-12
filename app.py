import io
import re
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from copy import copy

st.set_page_config(page_title="Nike Template Builder", layout="wide")

st.title("Nike Template Builder")

tpl_file = st.file_uploader("Carica TEMPLATE NIKE.xlsx", type=["xlsx"])
data_file = st.file_uploader("Carica file dati (CO_DF_1202 FILE.xlsx)", type=["xlsx"])

write_zeros = st.checkbox("Scrivi anche gli 0 nelle celle taglia", value=False)
start_row = st.number_input("Riga di partenza (dopo header)", min_value=2, value=2)

# =========================
# UTIL
# =========================

def normalize_size(x):
    if pd.isna(x):
        return ""
    s = str(x).strip().replace(",", ".")
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    if re.fullmatch(r"\d+\.\d+", s):
        s = str(float(s)).rstrip("0").rstrip(".")
    return s

def copy_row_style(ws, source_row, target_row, max_col):
    for c in range(1, max_col + 1):
        src = ws.cell(source_row, c)
        dst = ws.cell(target_row, c)
        if src.has_style:
            dst._style = copy(src._style)
        dst.number_format = src.number_format
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.alignment = copy(src.alignment)
        dst.protection = copy(src.protection)

def clear_row(ws, row, max_col):
    for c in range(1, max_col + 1):
        ws.cell(row, c).value = None


# =========================
# MAIN
# =========================

if tpl_file and data_file:

    df = pd.read_excel(data_file)

    required_cols = {"index", "size", "qty"}
    if not required_cols.issubset(df.columns):
        st.error("Il file dati deve contenere: index, size, qty")
        st.stop()

    df["size_norm"] = df["size"].apply(normalize_size)

    pivot = (
        df.pivot_table(
            index="index",
            columns="size_norm",
            values="qty",
            aggfunc="sum",
            fill_value=0
        )
        .sort_index()
    )

    wb = openpyxl.load_workbook(io.BytesIO(tpl_file.getvalue()))
    ws = wb.active

    # Trova header Material Number
    header_row = None
    for r in range(1, 30):
        if "Material Number" in [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]:
            header_row = r
            break

    if not header_row:
        st.error("Non trovo la riga con 'Material Number'")
        st.stop()

    headers = [ws.cell(header_row, c).value for c in range(1, ws.max_column + 1)]

    material_col = None
    for i, h in enumerate(headers, start=1):
        if str(h).strip() == "Material Number":
            material_col = i
            break

    if not material_col:
        st.error("Colonna 'Material Number' non trovata")
        st.stop()

    # =========================
    # RANGE TAGLIE FISSO JU → MP
    # =========================

    col_start = column_index_from_string("JU")
    col_end = column_index_from_string("MP")

    size_to_col = {}

    for col in range(col_start, col_end + 1):
        header_value = ws.cell(header_row, col).value
        if header_value is None:
            continue

        size_key = normalize_size(header_value)
        if size_key:
            size_to_col[size_key] = col

    # =========================
    # SCRITTURA
    # =========================

    if start_row <= header_row:
        start_row = header_row + 1

    max_col = ws.max_column
    style_source = start_row

    skus = list(pivot.index)

    for i, sku in enumerate(skus):
        r = start_row + i

        copy_row_style(ws, style_source, r, max_col)
        clear_row(ws, r, max_col)

        ws.cell(r, material_col).value = str(sku).strip()

        row_vals = pivot.loc[sku]

        for size_key, qty in row_vals.items():
            if size_key not in size_to_col:
                continue

            if qty == 0 and not write_zeros:
                continue

            col = size_to_col[size_key]
            ws.cell(r, col).value = int(qty)

    # Output
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    st.success(f"Creati {len(skus)} SKU con pivot taglie.")
    st.download_button(
        "Scarica Excel risultante",
        data=output.getvalue(),
        file_name="NIKE_TEMPLATE_FILLED.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Carica entrambi i file.")

