import io
import os
import re
from copy import copy

import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

# =========================
# STREAMLIT
# =========================
st.set_page_config(page_title="Nike Template Builder", layout="wide")
st.title("Nike Template Builder")

# =========================
# CONFIG
# =========================
TEMPLATE_PATH = "TEMPLATE NIKE.xlsx"
SOLD_TO_VALUE = 342694
SHIP_TO_VALUE = 342861

SIZE_COL_START_LETTER = "AN"
SIZE_COL_END_LETTER = "MP"

# =========================
# UI
# =========================
data_file = st.file_uploader(
    "Carica file dati (xlsx) con colonne: index, size, qty",
    type=["xlsx"]
)

col1, col2 = st.columns(2)
with col1:
    write_zeros = st.checkbox("Scrivi anche gli 0 nelle celle taglia", value=False)
with col2:
    start_row = st.number_input("Riga di partenza (dopo header)", min_value=2, value=2)

# =========================
# FUNZIONI
# =========================

def clean_key(x):
    """Chiave robusta per matching taglie"""
    if x is None:
        return ""
    return str(x).strip().upper().replace(",", ".")

def find_header_row(ws, needle="Material Number", max_scan_rows=60):
    for r in range(1, max_scan_rows + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if needle in vals:
            return r
    return None

def locate_column(headers, name):
    for i, h in enumerate(headers, start=1):
        if str(h).strip() == name:
            return i
    return None

def load_template():
    if os.path.exists(TEMPLATE_PATH):
        return openpyxl.load_workbook(TEMPLATE_PATH)

    here = os.path.dirname(os.path.abspath(__file__))
    candidate = os.path.join(here, TEMPLATE_PATH)
    if os.path.exists(candidate):
        return openpyxl.load_workbook(candidate)

    raise FileNotFoundError(f"Template non trovato: {TEMPLATE_PATH}")

def ensure_rows_with_style(ws, template_row, start_row, n_rows, max_col):
    last_needed = start_row + n_rows - 1
    if last_needed <= ws.max_row:
        return

    rows_to_add = last_needed - ws.max_row
    insert_at = ws.max_row + 1
    ws.insert_rows(insert_at, amount=rows_to_add)

    src_height = ws.row_dimensions[template_row].height

    for r in range(insert_at, insert_at + rows_to_add):
        ws.row_dimensions[r].height = src_height
        for c in range(1, max_col + 1):
            src = ws.cell(template_row, c)
            dst = ws.cell(r, c)
            if src.has_style:
                dst._style = copy(src._style)
            dst.font = copy(src.font)
            dst.border = copy(src.border)
            dst.fill = copy(src.fill)
            dst.number_format = src.number_format
            dst.alignment = copy(src.alignment)
            dst.protection = copy(src.protection)

def clear_row_values(ws, r, sold_to_col, ship_to_col, material_col, size_start, size_end):
    ws.cell(r, sold_to_col).value = None
    ws.cell(r, ship_to_col).value = None
    ws.cell(r, material_col).value = None
    for c in range(size_start, size_end + 1):
        ws.cell(r, c).value = None

def hide_leading_trailing_empty(ws, header_row, size_start, size_end, pivot):
    ordered_cols = list(range(size_start, size_end + 1))
    totals = []

    for col in ordered_cols:
        key = clean_key(ws.cell(header_row, col).value)
        if key in pivot.columns:
            totals.append(pivot[key].sum())
        else:
            totals.append(0)

    idx = [i for i, t in enumerate(totals) if t > 0]
    if not idx:
        return

    first = idx[0]
    last = idx[-1]

    # nasconde solo esterni
    for i in range(0, first):
        if totals[i] == 0:
            ws.column_dimensions[get_column_letter(ordered_cols[i])].hidden = True

    for i in range(last + 1, len(ordered_cols)):
        if totals[i] == 0:
            ws.column_dimensions[get_column_letter(ordered_cols[i])].hidden = True

# =========================
# MAIN
# =========================
if not data_file:
    st.info("Carica il file dati.")
    st.stop()

df = pd.read_excel(data_file)

required = {"index", "size", "qty"}
if not required.issubset(df.columns):
    st.error(f"Il file deve contenere: {required}")
    st.stop()

df["size_key"] = df["size"].apply(clean_key)

pivot = (
    df.pivot_table(
        index="index",
        columns="size_key",
        values="qty",
        aggfunc="sum",
        fill_value=0
    )
)

pivot.columns = [clean_key(c) for c in pivot.columns]

wb = load_template()
ws = wb.active

header_row = find_header_row(ws)
headers = [ws.cell(header_row, c).value for c in range(1, ws.max_column + 1)]

material_col = locate_column(headers, "Material Number")
sold_to_col = locate_column(headers, "Sold To")
ship_to_col = locate_column(headers, "Ship To")

size_start = column_index_from_string(SIZE_COL_START_LETTER)
size_end = column_index_from_string(SIZE_COL_END_LETTER)

start_row = int(start_row)
if start_row <= header_row:
    start_row = header_row + 1

skus = list(pivot.index)
max_col = ws.max_column

template_row_for_style = start_row if start_row <= ws.max_row else header_row + 1

ensure_rows_with_style(ws, template_row_for_style, start_row, len(skus), max_col)

# crea mapping template
key_to_col = {}
for col in range(size_start, size_end + 1):
    key = clean_key(ws.cell(header_row, col).value)
    if key:
        key_to_col[key] = col

# scrittura dati
for i, sku in enumerate(skus):
    r = start_row + i

    clear_row_values(ws, r, sold_to_col, ship_to_col, material_col, size_start, size_end)

    ws.cell(r, sold_to_col).value = SOLD_TO_VALUE
    ws.cell(r, ship_to_col).value = SHIP_TO_VALUE
    ws.cell(r, material_col).value = str(sku)

    for size_key, qty in pivot.loc[sku].items():
        if size_key in key_to_col:
            if qty == 0 and not write_zeros:
                continue
            ws.cell(r, key_to_col[size_key]).value = int(qty)

# nasconde solo esterni
hide_leading_trailing_empty(ws, header_row, size_start, size_end, pivot)

# output
out = io.BytesIO()
wb.save(out)
out.seek(0)

st.success(f"Creato file con {len(skus)} SKU.")
st.download_button(
    "Scarica Excel",
    data=out.getvalue(),
    file_name="NIKE_TEMPLATE_FILLED.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
