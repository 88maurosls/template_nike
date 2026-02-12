import io
import os
import re

import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.utils import column_index_from_string

st.set_page_config(page_title="Nike Template Builder", layout="wide")
st.title("Nike Template Builder")

# =========================
# CONFIG
# =========================
TEMPLATE_PATH = "TEMPLATE NIKE.xlsx"   # nella repo
SOLD_TO_VALUE = 342694
SHIP_TO_VALUE = 342861

SIZE_COL_START_LETTER = "JU"
SIZE_COL_END_LETTER = "MP"

data_file = st.file_uploader("Carica file dati (xlsx) con colonne: index, size, qty", type=["xlsx"])

col1, col2 = st.columns([1, 1])
with col1:
    write_zeros = st.checkbox("Scrivi anche gli 0 nelle celle taglia", value=False)
with col2:
    start_row = st.number_input("Riga di partenza (dopo header)", min_value=2, value=2, step=1)

# =========================
# UTILS
# =========================
def normalize_size(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip().replace(",", ".")
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    if re.fullmatch(r"\d+\.\d+", s):
        s = str(float(s)).rstrip("0").rstrip(".")
    return s

def find_header_row(ws, needle="Material Number", max_scan_rows=60):
    for r in range(1, max_scan_rows + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if needle in vals:
            return r
    return None

def locate_column(headers, name: str):
    for i, h in enumerate(headers, start=1):
        if str(h).strip() == name:
            return i
    return None

def load_template_workbook():
    if os.path.exists(TEMPLATE_PATH):
        return openpyxl.load_workbook(TEMPLATE_PATH)

    here = os.path.dirname(os.path.abspath(__file__))
    candidate = os.path.join(here, TEMPLATE_PATH)
    if os.path.exists(candidate):
        return openpyxl.load_workbook(candidate)

    raise FileNotFoundError(
        f"Template non trovato. Atteso '{TEMPLATE_PATH}' nella stessa cartella di app.py."
    )

def clear_only_needed_cells(ws, r, sold_to_col, ship_to_col, material_col, size_col_start, size_col_end):
    # NON tocchiamo stili: solo valori
    ws.cell(r, sold_to_col).value = None
    ws.cell(r, ship_to_col).value = None
    ws.cell(r, material_col).value = None
    for c in range(size_col_start, size_col_end + 1):
        ws.cell(r, c).value = None

# =========================
# MAIN
# =========================
if not data_file:
    st.info("Carica il file dati per generare l'Excel.")
    st.stop()

df = pd.read_excel(data_file)

required_cols = {"index", "size", "qty"}
if not required_cols.issubset(df.columns):
    st.error(f"Il file dati deve contenere le colonne: {sorted(list(required_cols))}. Trovate: {list(df.columns)}")
    st.stop()

df = df.copy()
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

try:
    wb = load_template_workbook()
except Exception as e:
    st.error(str(e))
    st.stop()

ws = wb.active

header_row = find_header_row(ws, "Material Number")
if not header_row:
    st.error("Non trovo la riga header con 'Material Number' nel template.")
    st.stop()

headers = [ws.cell(header_row, c).value for c in range(1, ws.max_column + 1)]

material_col = locate_column(headers, "Material Number")
sold_to_col = locate_column(headers, "Sold To")
ship_to_col = locate_column(headers, "Ship To")

if not material_col:
    st.error("Colonna 'Material Number' non trovata nel template.")
    st.stop()
if not sold_to_col or not ship_to_col:
    st.error("Colonne 'Sold To' e/o 'Ship To' non trovate nel template.")
    st.stop()

size_col_start = column_index_from_string(SIZE_COL_START_LETTER)
size_col_end = column_index_from_string(SIZE_COL_END_LETTER)

# mappa taglia -> colonna solo in JU:MP
size_to_col = {}
for col in range(size_col_start, size_col_end + 1):
    hv = ws.cell(header_row, col).value
    if hv is None:
        continue
    key = normalize_size(hv)
    if key:
        size_to_col[key] = col

start_row = int(start_row)
if start_row <= header_row:
    start_row = header_row + 1

skus = list(pivot.index)
last_needed_row = start_row + len(skus) - 1

# 🔒 REGOLA: NON inseriamo righe, per non perdere stile tabella/layout
if last_needed_row > ws.max_row:
    st.error(
        f"Il template non ha abbastanza righe formattate.\n\n"
        f"Servono almeno fino alla riga {last_needed_row}, ma il foglio arriva a {ws.max_row}.\n"
        f"Aggiungi nel TEMPLATE righe vuote già formattate (copiando una riga esistente) e riprova."
    )
    st.stop()

missing_sizes = [s for s in pivot.columns if str(s).strip() and str(s).strip() not in size_to_col]
if missing_sizes:
    st.warning("Taglie nel file dati non trovate tra JU e MP nel template: " + ", ".join(map(str, missing_sizes)))

# scrivi dentro righe esistenti (layout preservato)
for i, sku in enumerate(skus):
    r = start_row + i

    clear_only_needed_cells(
        ws, r,
        sold_to_col=sold_to_col,
        ship_to_col=ship_to_col,
        material_col=material_col,
        size_col_start=size_col_start,
        size_col_end=size_col_end
    )

    ws.cell(r, sold_to_col).value = SOLD_TO_VALUE
    ws.cell(r, ship_to_col).value = SHIP_TO_VALUE
    ws.cell(r, material_col).value = str(sku).strip()

    row_vals = pivot.loc[sku]
    for size_key, qty in row_vals.items():
        size_key = str(size_key).strip()
        if not size_key or size_key not in size_to_col:
            continue
        if qty == 0 and not write_zeros:
            continue
        ws.cell(r, size_to_col[size_key]).value = int(qty)

out = io.BytesIO()
wb.save(out)
out.seek(0)

st.success(f"Creato file con {len(skus)} SKU. Layout del template preservato (nessuna riga inserita).")
st.download_button(
    "Scarica Excel risultante",
    data=out.getvalue(),
    file_name="NIKE_TEMPLATE_FILLED.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
