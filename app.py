import io, json, os
from pathlib import Path
import pandas as pd
from docx import Document
from docx.shared import Pt
import streamlit as st

# ------------------------------
# ---- VERSION TRACKING ----
# ------------------------------
DATA_DIR = Path(".data")
DATA_DIR.mkdir(exist_ok=True)
VERSION_STORE = DATA_DIR / "version_store.json"
if VERSION_STORE.exists():
    version_tracker = json.loads(VERSION_STORE.read_text())
else:
    version_tracker = {}

# ------------------------------
# ---- CORE LOGIC PLACEHOLDER ----
# ------------------------------
st.title("PM Agreement Scope Generator")
st.caption("Upload your completed Service Intake Excel to generate a formatted DOCX scope.")

uploaded = st.file_uploader("Upload Service Intake (.xlsx)", type=["xlsx"])
prop_name = st.text_input("Property Name (for filename)", value="")
alpha = st.checkbox("Alphabetize equipment sections", value=True)

if st.button("Generate Scope") and uploaded is not None:
    st.success("Scope would be generated here (logic coming soon).")
