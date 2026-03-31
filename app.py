import io
import re
from datetime import datetime

import streamlit as st
import pdfplumber
import fitz  # PyMuPDF
from openpyxl import load_workbook
from openpyxl.styles import Font

st.set_page_config(page_title="BH Sanctions List Updater", page_icon="🛡️", layout="centered")

COLUMNS = [
    "NAME", "AKA", "FOREIGN_SCRIPT", "SEX", "DOB", "POB",
    "NATIONALITY", "OTHER_INFO", "ADD", "ADD_COUNTRY",
    "TITLE", "CITIZENSHIP", "REMARK",
]

# ---------------- PDF TEXT EXTRACTION ----------------

def extract_text_pymupdf(pdf_bytes: bytes) -> str:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    parts = []
    for page in doc:
        t = page.get_text("text")  # robust text extraction
        if t:
            parts.append(t)
    return "\n".join(parts).strip()

def extract_text_pdfplumber(pdf_bytes: bytes) -> str:
    parts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                parts.append(t)
    return "\n".join(parts).strip()

def extract_pdf_text(pdf_bytes: bytes) -> tuple[str, str]:
    """
    Returns (text, method_used)
    """
    text = extract_text_pymupdf(pdf_bytes)
    if text:
        return text, "PyMuPDF"
    text = extract_text_pdfplumber(pdf_bytes)
    if text:
        return text, "pdfplumber"
    return "", "none"

# ---------------- NAME PARSING ----------------

NAME_LINE = re.compile(r"(?im)^\s*Name\s*:\s*(.+?)\s*$")
NAME_ORIG = re.compile(r"(?im)^\s*Name\s*\(original.*?\)\s*:\s*(.+?)\s*$")
AKA_LINE  = re.compile(r"(?im)^\s*(A\.k\.a|AKA|Aliases?)\s*:\s*(.+?)\s*$")
DOB_LINE  = re.compile(r"(?im)^\s*(DOB|Date of birth)\s*:\s*(.+?)\s*$")
NAT_LINE  = re.compile(r"(?im)^\s*(Nationality)\s*:\s*(.+?)\s*$")

def normalize_space(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()

def extract_individuals_from_text(text: str) -> list[dict]:
    """
    Extracts entries using:
      1) Name: ...
      2) Name (original script): ...
    Then groups nearby fields (AKA/DOB/Nationality) in a small window.
    """
    lines = text.splitlines()
    individuals = []

    # Gather indices where "Name:" appears
    name_hits = []
    for i, line in enumerate(lines):
        m = NAME_LINE.match(line)
        if m:
            name_hits.append((i, normalize_space(m.group(1))))

    # If no "Name:" hits, fallback: detect likely English-name-ish lines
    # (Not ALL CAPS only; allow Title Case / hyphens / apostrophes)
    if not name_hits:
        likely = re.compile(r"^[A-Za-z][A-Za-z'.-]+(?:\s+[A-Za-z][A-Za-z'.-]+){1,6}$")
        for i, line in enumerate(lines):
            s = normalize_space(line)
            if likely.match(s) and len(s) >= 8:
                name_hits.append((i, s))

    # Build records using local window search around each hit
    for idx, nm in name_hits:
        window = "\n".join(lines[idx: idx + 25])  # look ahead 25 lines
        foreign = ""
        aka = ""
        dob = ""
        nat = ""

        mo = NAME_ORIG.search(window)
        if mo:
            foreign = normalize_space(mo.group(1))

        ma = AKA_LINE.search(window)
        if ma:
            aka = normalize_space(ma.group(2))

        md = DOB_LINE.search(window)
        if md:
            dob = normalize_space(md.group(2))

        mn = NAT_LINE.search(window)
        if mn:
            nat = normalize_space(mn.group(2))

        individuals.append({
            "NAME": nm,
            "AKA": aka,
            "FOREIGN_SCRIPT": foreign,
            "SEX": "",
            "DOB": dob,
            "POB": "",
            "NATIONALITY": nat,
            "OTHER_INFO": "Auto-extracted from Gazette PDF (offline parser)",
            "ADD": "",
            "ADD_COUNTRY": "",
            "TITLE": "",
            "CITIZENSHIP": "",
            "REMARK": "",
        })

    # De-duplicate by NAME (case-insensitive)
    seen = set()
    uniq = []
    for r in individuals:
        k = r["NAME"].strip().lower()
        if k and k not in seen:
            seen.add(k)
            uniq.append(r)

    return uniq

# ---------------- XLSX UPDATE ----------------

def update_xlsx(xlsx_file, individuals: list[dict]) -> tuple[bytes, int, int]:
    wb = load_workbook(xlsx_file)
    ws = wb.active

    existing = {
        str(row[0]).strip().lower()
        for row in ws.iter_rows(min_row=2, values_only=True)
        if row and row[0]
    }

    added = skipped = 0

    for person in individuals:
        name = (person.get("NAME") or "").strip()
        key = name.lower()

        if not key or key in existing:
            skipped += 1
            continue

        row_num = ws.max_row + 1
        for ci, col_name in enumerate(COLUMNS, start=1):
            val = person.get(col_name) or None
            cell = ws.cell(row=row_num, column=ci, value=val)
            cell.font = Font(name="Calibri", size=11)

        existing.add(key)
        added += 1

    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read(), added, skipped


# ---------------- UI ----------------

st.markdown("## 🛡️ BH Sanctions List Updater (OFFLINE)")
st.caption("Upload Gazette PDF + BH‑TL‑INDIVIDUALS.xlsx → download updated XLSX. No API. No JSON.")
st.divider()

c1, c2 = st.columns(2)
with c1:
    pdf_file = st.file_uploader("📄 Upload Gazette PDF", type=["pdf"])
with c2:
    xlsx_file = st.file_uploader("📊 Upload BH‑TL‑INDIVIDUALS.xlsx", type=["xlsx", "xls"])

run = st.button("▶ Extract & Update XLSX", type="primary", disabled=not (pdf_file and xlsx_file))

if run:
    pdf_bytes = pdf_file.read()
    text, method = extract_pdf_text(pdf_bytes)

    # Diagnostics first
    with st.expander("🧪 Diagnostics (important)", expanded=True):
        st.write(f"Extractor used: **{method}**")
        st.write(f"Extracted text length: **{len(text)}** characters")
        st.code((text[:1500] + ("…" if len(text) > 1500 else "")) or "[EMPTY TEXT]", language="text")

        if not text:
            st.error(
                "No extractable text was found. This usually means the PDF is a scanned image. "
                "Offline text extraction won’t work unless you add OCR."
            )

    if not text:
        st.stop()

    individuals = extract_individuals_from_text(text)

    st.info(f"Found **{len(individuals)}** unique candidate name(s).")

    with st.expander("👀 Preview extracted names", expanded=False):
        for i, r in enumerate(individuals[:50], 1):
            st.write(f"{i}. {r['NAME']}")
        if len(individuals) > 50:
            st.caption(f"Showing first 50 of {len(individuals)}")

    out_bytes, added, skipped = update_xlsx(xlsx_file, individuals)

    st.success(f"Done — Added: **{added}**, Skipped (duplicates/empty): **{skipped}**")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = f"{xlsx_file.name.rsplit('.',1)[0]}-UPDATED-{ts}.xlsx"
    st.download_button(
        "⬇️ Download Updated XLSX",
        data=out_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
