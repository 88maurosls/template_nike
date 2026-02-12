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
        if nee
