import io
import os
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
SOLD_TO_VALUE = 000000
SHIP_TO_VALUE = 000000

SIZE_COL_START_LETTER = "AM"
SIZE_COL_END_LETTER = "DR"

# =========================
# UI
# =========================
data_file = st.file_uploader(
    "Carica file dati (xlsx) con colonne: index, size, qty",
    type=["xlsx"]
)

# =========================
# FUNZIONI
# =========================

def clean_key(x):
    if x is None:
        return ""
    return str(x).strip().upper().replace(",", ".")

def find_header_row(ws, needle="Material Number", max_scan_rows=50):
    for r in range(1, max_scan_rows + 1):
        row_values = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if needle in row_values:
            return r
    return None

def ensure_rows_with_style(ws, template_row, start_row, n_rows, max_col):
    """
    Se servono più righe di quelle presenti, le crea e copia lo stile dal template_row.
    """
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

def set_bold(cell):
    """
    Forza bold sulla cella mantenendo il resto (font name, size, ecc).
    """
    f = copy(cell.font)
    f.bold = True
    cell.font = f

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
        aggfunc="sum"
    )
)

pivot.columns = [clean_key(c) for c in pivot.columns]

if not os.path.exists(TEMPLATE_PATH):
    st.error("Template non trovato.")
    st.stop()

wb = openpyxl.load_workbook(TEMPLATE_PATH)
ws = wb.active

header_row = find_header_row(ws)
if not header_row:
    st.error("Header non trovato nel template.")
    st.stop()

headers = [ws.cell(header_row, c).value for c in range(1, ws.max_column + 1)]

material_col = headers.index("Material Number") + 1
sold_to_col = headers.index("Sold To") + 1
ship_to_col = headers.index("Ship To") + 1

size_start = column_index_from_string(SIZE_COL_START_LETTER)
size_end = column_index_from_string(SIZE_COL_END_LETTER)

start_row = header_row + 1
template_style_row = start_row  # riga “modello” nel template

# assicura righe sufficienti e copia stile
max_col = ws.max_column
ensure_rows_with_style(ws, template_style_row, start_row, len(pivot.index), max_col)

# mapping taglie
key_to_col = {}
for col in range(size_start, size_end + 1):
    key = clean_key(ws.cell(header_row, col).value)
    if key:
        key_to_col[key] = col

# scrittura dati
for i, sku in enumerate(pivot.index):
    r = start_row + i

    c = ws.cell(r, sold_to_col)
    c.value = SOLD_TO_VALUE
    set_bold(c)

    c = ws.cell(r, ship_to_col)
    c.value = SHIP_TO_VALUE
    set_bold(c)

    c = ws.cell(r, material_col)
    c.value = str(sku)
    set_bold(c)

    for size_key, qty in pivot.loc[sku].items():
        if size_key in key_to_col and pd.notna(qty) and qty > 0:
            c = ws.cell(r, key_to_col[size_key])
            c.value = int(qty)
            set_bold(c)

# forza visibilità colonne taglie AM -> DR
for col in range(size_start, size_end + 1):
    ws.column_dimensions[get_column_letter(col)].hidden = False

# output
out = io.BytesIO()
wb.save(out)
out.seek(0)

st.success(f"Creato file con {len(pivot.index)} SKU.")
st.download_button(
    "Scarica Excel",
    data=out.getvalue(),
    file_name="NIKE_TEMPLATE_FILLED.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
