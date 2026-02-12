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
TEMPLATE_PATH = "TEMPLATE NIKE.xlsx"   # nella repo (stessa cartella di app.py)
SOLD_TO_VALUE = 342694
SHIP_TO_VALUE = 342861

# RANGE TAGLIE: AN -> MP
SIZE_COL_START_LETTER = "AN"
SIZE_COL_END_LETTER = "MP"

# =========================
# UI
# =========================
data_file = st.file_uploader("Carica file dati (xlsx) con colonne: index, size, qty", type=["xlsx"])

col1, col2 = st.columns([1, 1])
with col1:
    write_zeros = st.checkbox("Scrivi anche gli 0 nelle celle taglia", value=False)
with col2:
    start_row = st.number_input("Riga di partenza (dopo header)", min_value=2, value=2, step=1)

# =========================
# FUNZIONI
# =========================
def normalize_size(x) -> str:
    """
    Normalizza taglie numeriche: 8.0->8, 8.50->8.5
    Se non è numerico, ritorna stringa pulita (es: S, M, 1C, 1.5C, ecc).
    """
    if pd.isna(x):
        return ""

    s = str(x).strip()
    if not s:
        return ""

    s = s.replace(",", ".").upper()

    # se è un numero puro
    if re.fullmatch(r"\d+(\.\d+)?", s):
        if re.fullmatch(r"\d+\.0", s):
            s = s[:-2]
        if re.fullmatch(r"\d+\.\d+", s):
            s = str(float(s)).rstrip("0").rstrip(".")
        return s

    # taglie tipo 1C, 1.5C, 2C, ecc
    if re.fullmatch(r"\d+(\.\d+)?C", s):
        # normalizza la parte numerica
        num = s[:-1]
        if re.fullmatch(r"\d+\.0", num):
            num = num[:-2]
        if re.fullmatch(r"\d+\.\d+", num):
            num = str(float(num)).rstrip("0").rstrip(".")
        return f"{num}C"

    # letterali (S, M, L, XL, ecc) o altro: lascio così
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

def ensure_rows_with_style(ws, template_row, start_row, n_rows, max_col):
    """
    Garantisce che esistano n_rows righe a partire da start_row.
    Se mancano, inserisce righe e copia stile COMPLETO + altezza dalla template_row.
    """
    last_needed = start_row + n_rows - 1
    if last_needed <= ws.max_row:
        return

    rows_to_add = last_needed - ws.max_row
    insert_at = ws.max_row + 1
    ws.insert_rows(insert_at, amount=rows_to_add)

    src_h = ws.row_dimensions[template_row].height
    for r in range(insert_at, insert_at + rows_to_add):
        ws.row_dimensions[r].height = src_h
        for c in range(1, max_col + 1):
            src = ws.cell(template_row, c)
            dst = ws.cell(r, c)

            if src.has_style:
                dst._style = copy(src._style)
            dst.number_format = src.number_format
            dst.font = copy(src.font)
            dst.border = copy(src.border)
            dst.fill = copy(src.fill)
            dst.alignment = copy(src.alignment)
            dst.protection = copy(src.protection)

def clear_only_needed_cells(ws, r, sold_to_col, ship_to_col, material_col, size_col_start, size_col_end):
    """
    Pulisce SOLO i valori delle celle gestite.
    """
    ws.cell(r, sold_to_col).value = None
    ws.cell(r, ship_to_col).value = None
    ws.cell(r, material_col).value = None
    for c in range(size_col_start, size_col_end + 1):
        ws.cell(r, c).value = None

def hide_leading_trailing_empty_size_cols(ws, header_row, size_col_start, size_col_end, pivot):
    """
    Nasconde SOLO colonne vuote prima e dopo, senza modificare le altre.
    Non fa mai hidden=False (così non cambia il template).
    Vuoto = somma qty == 0 per quella taglia.
    Le colonne "vuote in mezzo" restano visibili.
    """
    ordered_cols = list(range(size_col_start, size_col_end + 1))

    totals = []
    for col in ordered_cols:
        hv = ws.cell(header_row, col).value
        key = normalize_size(hv)

        # se header vuoto, la consideriamo "vuota"
        if not key:
            totals.append(0)
            continue

        if key in pivot.columns:
            # pivot columns sono stringhe normalizzate
            totals.append(float(pivot[key].sum()))
        else:
            totals.append(0)

    # trova prima e ultima colonna con totale > 0
    idxs = [i for i, t in enumerate(totals) if t > 0]
    if not idxs:
        return  # nulla da fare

    first = idxs[0]
    last = idxs[-1]

    # nascondi SOLO prima
    for i in range(0, first):
        if totals[i] == 0:
            letter = get_column_letter(ordered_cols[i])
            ws.column_dimensions[letter].hidden = True

    # nascondi SOLO dopo
    for i in range(last + 1, len(ordered_cols)):
        if totals[i] == 0:
            letter = get_column_letter(ordered_cols[i])
            ws.column_dimensions[letter].hidden = True

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

# range taglie fisso AN -> MP
size_col_start = column_index_from_string(SIZE_COL_START_LETTER)
size_col_end = column_index_from_string(SIZE_COL_END_LETTER)

# start row effettiva
start_row = int(start_row)
if start_row <= header_row:
    start_row = header_row + 1

skus = list(pivot.index)
max_col = ws.max_column

template_row_for_style = start_row if start_row <= ws.max_row else header_row + 1

ensure_rows_with_style(
    ws=ws,
    template_row=template_row_for_style,
    start_row=start_row,
    n_rows=len(skus),
    max_col=max_col
)

# warning taglie nel pivot non presenti nel template (solo info)
# qui non mappiamo più via header->colonna perché scriviamo direttamente leggendo header
# ma ci serve comunque per avvisare l'utente
template_keys = set()
for col in range(size_col_start, size_col_end + 1):
    template_keys.add(normalize_size(ws.cell(header_row, col).value))
template_keys.discard("")

missing_sizes = [c for c in pivot.columns if str(c).strip() and str(c).strip() not in template_keys]
if missing_sizes:
    st.warning("Taglie nel file dati non trovate nel range AN-MP del template: " + ", ".join(map(str, missing_sizes)))

# crea mappa key -> col (prima occorrenza) nel range AN-MP
key_to_col = {}
for col in range(size_col_start, size_col_end + 1):
    key = normalize_size(ws.cell(header_row, col).value)
    if key and key not in key_to_col:
        key_to_col[key] = col

# scrittura righe
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
        if not size_key:
            continue
        if size_key not in key_to_col:
            continue
        if qty == 0 and not write_zeros:
            continue
        ws.cell(r, key_to_col[size_key]).value = int(qty)

# nascondi colonne vuote prima e dopo (senza mai fare hidden=False)
hide_leading_trailing_empty_size_cols(
    ws=ws,
    header_row=header_row,
    size_col_start=size_col_start,
    size_col_end=size_col_end,
    pivot=pivot
)

out = io.BytesIO()
wb.save(out)
out.seek(0)

st.success(f"Creato file con {len(skus)} SKU. Range taglie AN-MP e hide esterni (senza alterare il template).")
st.download_button(
    "Scarica Excel risultante",
    data=out.getvalue(),
    file_name="NIKE_TEMPLATE_FILLED.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
