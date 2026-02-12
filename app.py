import io
import os
import re

import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import range_boundaries, get_column_letter

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

def load_template_workbook():
    if os.path.exists(TEMPLATE_PATH):
        return openpyxl.load_workbook(TEMPLATE_PATH)
    here = os.path.dirname(os.path.abspath(__file__))
    candidate = os.path.join(here, TEMPLATE_PATH)
    if os.path.exists(candidate):
        return openpyxl.load_workbook(candidate)
    raise FileNotFoundError(f"Template non trovato: '{TEMPLATE_PATH}' (atteso nella stessa cartella di app.py).")

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

def clear_only_needed_cells(ws, r, sold_to_col, ship_to_col, material_col, size_col_start, size_col_end):
    # NON tocchiamo stili: solo valori
    ws.cell(r, sold_to_col).value = None
    ws.cell(r, ship_to_col).value = None
    ws.cell(r, material_col).value = None
    for c in range(size_col_start, size_col_end + 1):
        ws.cell(r, c).value = None

def expand_excel_table_to_row(ws, header_row, material_col, last_row_needed):
    """
    Se nel foglio esiste una Tabella Excel (Formato come tabella),
    espande il suo range fino a last_row_needed.
    Questo è ciò che fa tornare i colori/banding in Excel.
    """
    if not hasattr(ws, "tables") or len(ws.tables) == 0:
        return False, "Nessuna tabella trovata nel foglio."

    # scegli la tabella più plausibile: quella che include la cella di header "Material Number"
    chosen = None
    chosen_name = None

    for tname, table in ws.tables.items():
        # table.ref tipo "A1:LI10"
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        if min_row == header_row and (min_col <= material_col <= max_col):
            chosen = table
            chosen_name = tname
            break

    # fallback: prendi la prima tabella
    if chosen is None:
        chosen_name = list(ws.tables.keys())[0]
        chosen = ws.tables[chosen_name]

    min_col, min_row, max_col, max_row = range_boundaries(chosen.ref)

    # aggiorna ref solo se serve
    if last_row_needed > max_row:
        new_ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{last_row_needed}"
        chosen.ref = new_ref
        if getattr(chosen, "autoFilter", None) is not None:
            chosen.autoFilter.ref = new_ref
        return True, f"Tabella '{chosen_name}' espansa a {new_ref}."
    else:
        return True, f"Tabella '{chosen_name}' già copre fino a riga {max_row} (nessuna espansione necessaria)."

# =========================
# MAIN
# =========================
if not data_file:
    st.info("Carica il file dati per generare l'Excel.")
    st.stop()

df = pd.read_excel(data_file)

required_cols = {"index", "size", "qty"}
if not required_cols.issubset(df.columns):
    st.error(f"Il file dati deve contenere: {sorted(list(required_cols))}. Trovate: {list(df.columns)}")
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

# Qui non inseriamo righe (per non rompere nulla).
# Però ora espandiamo la TABELLA Excel, così i colori si applicano fino a last_needed_row.
ok_table, msg_table = expand_excel_table_to_row(ws, header_row, material_col, last_needed_row)
st.caption(msg_table)

# Controllo: il foglio deve comunque avere almeno quelle righe "fisicamente" (Excel ce le ha sempre),
# ma se il template è corto e Excel mostra solo fino a X righe, non è un problema.
# openpyxl può scrivere anche oltre ws.max_row senza inserire? No: deve esistere la riga.
# In pratica ws.max_row cresce quando assegni valori, quindi va bene.

missing_sizes = [s for s in pivot.columns if str(s).strip() and str(s).strip() not in size_to_col]
if missing_sizes:
    st.warning("Taglie nel file dati non trovate tra JU e MP nel template: " + ", ".join(map(str, missing_sizes)))

for i, sku in enumerate(skus):
    r = start_row + i

    # pulisci solo celle gestite
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

st.success("File creato usando il template e mantenendo i colori tramite espansione della Tabella.")
st.download_button(
    "Scarica Excel risultante",
    data=out.getvalue(),
    file_name="NIKE_TEMPLATE_FILLED.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
