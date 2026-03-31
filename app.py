import io
import re
from datetime import datetime

import pdfplumber
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font

# ── Config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="BH Sanctions List Updater",
    page_icon="🛡️",
    layout="centered",
)

COLUMNS = [
    "NAME", "AKA", "FOREIGN_SCRIPT", "SEX", "DOB", "POB",
    "NATIONALITY", "OTHER_INFO", "ADD", "ADD_COUNTRY",
    "TITLE", "CITIZENSHIP", "REMARK",
]

# ── Header ────────────────────────────────────────────────────────────────
st.markdown("## 🛡️ BH Sanctions List Updater (OFFLINE)")
st.caption(
    "Upload Bahrain Official Gazette PDF + BH‑TL‑INDIVIDUALS.xlsx → "
    "download updated XLSX. No API. No JSON."
)
st.divider()

# ── Uploaders ─────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.markdown("**📄 Official Gazette PDF**")
    pdf_file = st.file_uploader("Upload PDF", type=["pdf"], label_visibility="collapsed")

with col2:
    st.markdown("**📊 BH‑TL‑INDIVIDUALS.xlsx**")
    xlsx_file = st.file_uploader("Upload XLSX", type=["xlsx"], label_visibility="collapsed")

st.divider()

# ── Extract text from PDF ─────────────────────────────────────────────────
def extract_pdf_text(pdf_bytes) -> str:
    text = []
    with pdfplumber.open(pdf_bytes) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text.append(page_text)
    return "\n".join(text)

# ── Rule-based extraction (adjust regex if needed) ────────────────────────
def extract_individuals_from_text(text: str):
    individuals = []

    # Example pattern: English name lines (ALL CAPS typical in gazettes)
    name_pattern = re.compile(r"^[A-Z][A-Z\\s'.-]{6,}$", re.MULTILINE)

    matches = name_pattern.findall(text)

    for name in matches:
        individuals.append({
            "NAME": name.strip(),
            "AKA": "",
            "FOREIGN_SCRIPT": "",
            "SEX": "",
            "DOB": "",
            "POB": "",
            "NATIONALITY": "",
            "OTHER_INFO": "Extracted from Bahrain Official Gazette (auto)",
            "ADD": "",
            "ADD_COUNTRY": "",
            "TITLE": "",
            "CITIZENSHIP": "",
            "REMARK": "",
        })

    return individuals

# ── Run button ────────────────────────────────────────────────────────────
run = st.button(
    "▶ Extract & Update XLSX",
    disabled=not (pdf_file and xlsx_file),
    type="primary",
)

if not (pdf_file and xlsx_file):
    st.caption("⬆ Upload both PDF and XLSX to continue")

# ── Processing ────────────────────────────────────────────────────────────
if run:
    logs = []
    added = skipped = 0
    output_bytes = None
    error = None

    with st.spinner("Processing Gazette PDF..."):

        try:
            text = extract_pdf_text(pdf_file)
            individuals = extract_individuals_from_text(text)
            logs.append(f"PDF parsed — {len(individuals)} candidate names found")
        except Exception as e:
            error = f"PDF parsing failed: {e}"

        if not error:
            try:
                wb = load_workbook(xlsx_file)
                ws = wb.active
                existing = {
                    str(row[0]).strip().lower()
                    for row in ws.iter_rows(min_row=2, values_only=True)
                    if row and row[0]
                }
            except Exception as e:
                error = f"XLSX load failed: {e}"

        if not error:
            for person in individuals:
                name = person["NAME"]
                key = name.lower()

                if key in existing:
                    skipped += 1
                    continue

                row_num = ws.max_row + 1
                for ci, col in enumerate(COLUMNS, start=1):
                    ws.cell(row=row_num, column=ci, value=person.get(col) or None).font = Font(size=11)

                existing.add(key)
                added += 1

            ws.freeze_panes = "A2"

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            output_bytes = buf.read()

    # ── Results ───────────────────────────────────────────────────────────
    if error:
        st.error(error)
    else:
        st.success("✅ Update complete")
        st.markdown(
            f"""
**Results**
- ✅ Added: **{added}**
- ⚠ Skipped (duplicates): **{skipped}**
"""
        )

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"BH-TL-INDIVIDUALS-UPDATED-{ts}.xlsx"

        st.download_button(
            "⬇️ Download Updated XLSX",
            data=output_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        with st.expander("📋 Extraction log"):
            for l in logs:
                st.markdown(f"• {l}")
