"""
BH Sanctions List Updater — Streamlit App (Offline / No API Key)
=================================================================
Deploy on Streamlit Cloud:
    1. Push this file + requirements.txt to GitHub
    2. Go to share.streamlit.io → deploy → done

requirements.txt:
    streamlit
    openpyxl
"""

import io
import json
from datetime import datetime

import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font

# ── Page config ────────────────────────────────────────────────────────────
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

# ── Custom CSS ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 2rem; max-width: 780px; }
    .title-badge {
        display: inline-block;
        background: #e8f4ff;
        color: #0066cc;
        border: 1px solid #b3d4f5;
        border-radius: 4px;
        padding: 2px 8px;
        font-size: 11px;
        font-weight: 700;
        margin-left: 8px;
        vertical-align: middle;
    }
    .result-row {
        display: flex;
        gap: 16px;
        margin-top: 12px;
        margin-bottom: 16px;
    }
    .metric-box {
        background: #f0faf5;
        border: 1px solid #b3e6cf;
        border-radius: 8px;
        padding: 14px 20px;
        text-align: center;
        flex: 1;
    }
    .metric-num { font-size: 30px; font-weight: 800; color: #00875a; }
    .metric-lbl { font-size: 11px; color: #5a6a7a; margin-top: 2px; }
    .metric-box.skip { background: #fffaf0; border-color: #f5d07a; }
    .metric-box.skip .metric-num { color: #b07d00; }
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────
st.markdown(
    '## 🛡️ BH Sanctions List Updater <span class="title-badge">OFFLINE</span>',
    unsafe_allow_html=True,
)
st.caption(
    "Upload the Gazette PDF + BH-TL-INDIVIDUALS.xlsx + extracted individuals JSON → "
    "download the updated XLSX. No API key required."
)
st.divider()

# ── How it works ───────────────────────────────────────────────────────────
with st.expander("ℹ️ How to prepare the JSON file", expanded=False):
    st.markdown("""
1. Upload the Gazette PDF to **Claude AI** (claude.ai) and ask:
   > *"Extract all sanctioned individuals from this PDF as a JSON array with keys:
   NAME, AKA, FOREIGN_SCRIPT, SEX, DOB, POB, NATIONALITY, OTHER_INFO,
   ADD, ADD_COUNTRY, TITLE, CITIZENSHIP, REMARK"*
2. Copy Claude's JSON response and save it as a `.json` file
3. Upload it here along with the PDF and XLSX
""")
    st.code('''[
  {
    "NAME": "Sami Jasim Muhammad Jaata Al-Jaburi",
    "AKA": "Mustafa Adnan al-Aziz; Sami al-Ajuz; Hajji Hamid",
    "FOREIGN_SCRIPT": "سامي جاسم محمد جعطة الجبوري",
    "SEX": "Male",
    "DOB": "1 Jul. 1974",
    "POB": "Iraq",
    "NATIONALITY": "Iraqi",
    "OTHER_INFO": "QDi.437. Senior ISIS commander...",
    "ADD": "Iraq",
    "ADD_COUNTRY": "Iraq",
    "TITLE": "",
    "CITIZENSHIP": "",
    "REMARK": "Gazette No. 3869, 29 March 2026"
  }
]''', language="json")

st.divider()

# ── File uploaders ─────────────────────────────────────────────────────────
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("**📄 Official Gazette PDF**")
    pdf_file = st.file_uploader(
        "Upload PDF", type=["pdf"],
        label_visibility="collapsed", key="pdf"
    )
    if pdf_file:
        st.success(f"✓ {pdf_file.name}")

with col2:
    st.markdown("**📊 BH-TL-INDIVIDUALS.xlsx**")
    xlsx_file = st.file_uploader(
        "Upload XLSX", type=["xlsx", "xls"],
        label_visibility="collapsed", key="xlsx"
    )
    if xlsx_file:
        st.success(f"✓ {xlsx_file.name}")

with col3:
    st.markdown("**{ } Individuals JSON**")
    json_file = st.file_uploader(
        "Upload JSON", type=["json"],
        label_visibility="collapsed", key="json"
    )
    if json_file:
        st.success(f"✓ {json_file.name}")

st.divider()

# ── Run button ─────────────────────────────────────────────────────────────
all_ready = pdf_file and xlsx_file and json_file

run = st.button(
    "▶  Update XLSX",
    disabled=not all_ready,
    type="primary",
)

if not all_ready:
    st.caption("⬆ Upload all three files above to enable processing.")

# ── Processing logic ───────────────────────────────────────────────────────
if run and all_ready:
    logs = []
    added = skipped = 0
    output_bytes = None
    error = None

    with st.spinner("Processing files..."):

        # Step 1 — Parse JSON
        try:
            raw = json_file.read().decode("utf-8")
            individuals = json.loads(raw)
            if not isinstance(individuals, list):
                raise ValueError("JSON must be an array of objects")
            individuals = [
                {k: (row.get(k, "") or "") for k in COLUMNS}
                for row in individuals
                if isinstance(row, dict)
            ]
            logs.append(("ok", f"JSON loaded — {len(individuals)} individual(s) found"))
        except Exception as e:
            error = f"Invalid JSON file: {e}"

        # Step 2 — Load XLSX
        if not error:
            try:
                wb = load_workbook(xlsx_file)
                ws = wb.active
                existing = {
                    str(row[0]).strip().lower()
                    for row in ws.iter_rows(min_row=2, values_only=True)
                    if row and row[0]
                }
                logs.append(("ok", f"XLSX loaded — {ws.max_row - 1} existing records"))
            except Exception as e:
                error = f"Cannot read XLSX: {e}"

        # Step 3 — Append new records
        if not error:
            for person in individuals:
                name = person.get("NAME", "").strip()
                key = name.lower()
                if not key or key in existing:
                    logs.append(("warn", f"Skipped duplicate: {name or '(empty name)'}"))
                    skipped += 1
                    continue
                row_num = ws.max_row + 1
                for ci, col_name in enumerate(COLUMNS, start=1):
                    val = person.get(col_name) or None
                    cell = ws.cell(row=row_num, column=ci, value=val)
                    cell.font = Font(name="Calibri", size=11)
                existing.add(key)
                logs.append(("ok", f"Added: {name}"))
                added += 1

            # Preserve original column widths
            for col_letter, width in {
                "A": 49.36, "B": 39.82, "C": 14.82,
                "E": 14.18, "F": 11.82, "G": 8.54,
                "H": 33.45, "K": 7.18, "L": 6.82, "M": 13.18,
            }.items():
                ws.column_dimensions[col_letter].width = width
            ws.freeze_panes = "A2"

            # Save to memory buffer
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            output_bytes = buf.read()
            logs.append(("ok", f"Done — {added} added, {skipped} skipped"))

    # ── Show processing log ────────────────────────────────────────────────
    with st.expander("📋 Processing log", expanded=True):
        for level, msg in logs:
            if level == "ok":
                st.markdown(f"✅ {msg}")
            elif level == "warn":
                st.markdown(f"⚠️ {msg}")
            else:
                st.markdown(f"ℹ️ {msg}")

    # ── Result ─────────────────────────────────────────────────────────────
    if error:
        st.error(f"**Error:** {error}")

    elif output_bytes:
        st.success("**Updated successfully!**")

        st.markdown(f"""
        <div class="result-row">
            <div class="metric-box">
                <div class="metric-num">{added}</div>
                <div class="metric-lbl">Individual(s) Added</div>
            </div>
            <div class="metric-box skip">
                <div class="metric-num">{skipped}</div>
                <div class="metric-lbl">Duplicate(s) Skipped</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        base = xlsx_file.name.rsplit(".", 1)[0]
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"{base}-UPDATED-{ts}.xlsx"

        st.download_button(
            label="⬇️  Download Updated XLSX",
            data=output_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
