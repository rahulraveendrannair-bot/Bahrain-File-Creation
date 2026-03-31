import io
import re
from datetime import datetime
from typing import List, Dict, Tuple, Optional

import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
import pytesseract
from PIL import Image
from openpyxl import load_workbook
from openpyxl.styles import Font


# =========================
# Configuration
# =========================
st.set_page_config(page_title="BH Sanctions List Updater", page_icon="🛡️", layout="centered")

COLUMNS = [
    "NAME", "AKA", "FOREIGN_SCRIPT", "SEX", "DOB", "POB",
    "NATIONALITY", "OTHER_INFO", "ADD", "ADD_COUNTRY",
    "TITLE", "CITIZENSHIP", "REMARK",
]

# Common junk / navigation strings that sometimes appear as the only "text"
GARBAGE_PATTERNS = [
    r"click here",
    r"individuals\s+click",
    r"entities\s+click",
    r"back to",
    r"home",
    r"download",
    r"http[s]?://",
    r"www\.",
    r"\btable of contents\b",
]

# Regex patterns inspired by typical Gazette/UN-style blocks that contain fields like Name/A.k.a/etc.
RE_NAME = re.compile(r"(?im)^\s*Name\s*:\s*(.+?)\s*$")
RE_NAME_ORIG = re.compile(r"(?im)^\s*Name\s*\(original.*?\)\s*:\s*(.+?)\s*$")
RE_AKA = re.compile(r"(?im)^\s*(A\.k\.a\.?|AKA|Aliases?)\s*:\s*(.+?)\s*$")
RE_DOB = re.compile(r"(?im)^\s*(DOB|Date of birth)\s*:\s*(.+?)\s*$")
RE_POB = re.compile(r"(?im)^\s*(POB|Place of birth)\s*:\s*(.+?)\s*$")
RE_NAT = re.compile(r"(?im)^\s*(Nationality)\s*:\s*(.+?)\s*$")
RE_ADDR = re.compile(r"(?im)^\s*(Address)\s*:\s*(.+?)\s*$")
RE_REMARK = re.compile(r"(?im)^\s*(Other information|Other info|Remarks?)\s*:\s*(.+?)\s*$")


# =========================
# Helpers
# =========================
def normalize_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def has_garbage_signature(text: str) -> bool:
    """
    True if extracted text is empty / too small / looks like navigation junk.
    """
    if not text or len(text.strip()) < 80:
        return True
    low = text.lower()
    # If a lot of matches are found, treat as junk
    hits = sum(1 for p in GARBAGE_PATTERNS if re.search(p, low))
    return hits >= 1 and len(text.strip()) < 800  # small + contains junk


# =========================
# PDF Text Extraction
# =========================
def extract_text_pymupdf_native(pdf_bytes: bytes) -> str:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    parts = []
    for page in doc:
        t = page.get_text("text")
        if t:
            parts.append(t)
    return "\n".join(parts).strip()


def extract_text_pdfplumber_native(pdf_bytes: bytes) -> str:
    parts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                parts.append(t)
    return "\n".join(parts).strip()


def extract_text_ocr_tesseract(pdf_bytes: bytes, dpi: int, lang: str) -> str:
    """
    OCR fallback for scanned/image-based PDFs:
      - Render pages to images with PyMuPDF
      - OCR each page with Tesseract via pytesseract
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)

    out = []
    for page in doc:
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        page_text = pytesseract.image_to_string(img, lang=lang)
        if page_text:
            out.append(page_text)

    return "\n".join(out).strip()


def extract_pdf_text_smart(pdf_bytes: bytes, dpi: int, lang: str) -> Tuple[str, str]:
    """
    Returns (text, method_used).
    Strategy:
      1) PyMuPDF native text
      2) pdfplumber native text
      3) If still empty/junk -> OCR (Tesseract)
    """
    text = extract_text_pymupdf_native(pdf_bytes)
    if text and not has_garbage_signature(text):
        return text, "PyMuPDF (native)"

    text2 = extract_text_pdfplumber_native(pdf_bytes)
    if text2 and not has_garbage_signature(text2):
        return text2, "pdfplumber (native)"

    # OCR fallback
    ocr_text = extract_text_ocr_tesseract(pdf_bytes, dpi=dpi, lang=lang)
    if ocr_text:
        return ocr_text, f"OCR (Tesseract) @ {dpi} DPI, lang={lang}"

    # nothing workable
    return "", "none"


# =========================
# Parsing Individuals
# =========================
def parse_individual_blocks(text: str) -> List[Dict[str, str]]:
    """
    Parse individuals using Name: ... blocks first (best signal).
    If no Name: matches, fall back to "likely name lines" but aggressively filter junk.
    """
    lines = text.splitlines()
    # Collect "Name:" hit indices
    name_hits: List[Tuple[int, str]] = []
    for i, line in enumerate(lines):
        m = RE_NAME.match(line)
        if m:
            candidate = normalize_space(m.group(1))
            if candidate and not any(re.search(p, candidate.lower()) for p in GARBAGE_PATTERNS):
                name_hits.append((i, candidate))

    individuals: List[Dict[str, str]] = []

    # If we have Name hits, treat each as start of a block
    if name_hits:
        for idx, nm in name_hits:
            window = "\n".join(lines[idx: idx + 30])  # look ahead within a local window

            foreign = ""
            aka = ""
            dob = ""
            pob = ""
            nat = ""
            addr = ""
            remark = ""

            mo = RE_NAME_ORIG.search(window)
            if mo:
                foreign = normalize_space(mo.group(1))

            ma = RE_AKA.search(window)
            if ma:
                aka = normalize_space(ma.group(2))

            md = RE_DOB.search(window)
            if md:
                dob = normalize_space(md.group(2))

            mp = RE_POB.search(window)
            if mp:
                pob = normalize_space(mp.group(2))

            mn = RE_NAT.search(window)
            if mn:
                nat = normalize_space(mn.group(2))

            mad = RE_ADDR.search(window)
            if mad:
                addr = normalize_space(mad.group(2))

            mr = RE_REMARK.search(window)
            if mr:
                remark = normalize_space(mr.group(2))

            individuals.append({
                "NAME": nm,
                "AKA": aka,
                "FOREIGN_SCRIPT": foreign,
                "SEX": "",
                "DOB": dob,
                "POB": pob,
                "NATIONALITY": nat,
                "OTHER_INFO": remark,
                "ADD": addr,
                "ADD_COUNTRY": "",
                "TITLE": "",
                "CITIZENSHIP": "",
                "REMARK": "",
            })

    else:
        # Fallback: attempt to detect likely name lines (Title Case / CAPS etc.)
        # BUT: filter out anything that looks like navigation / labels.
        likely = re.compile(r"^[A-Za-z][A-Za-z'.-]+(?:\s+[A-Za-z][A-Za-z'.-]+){1,6}$")

        for line in lines:
            s = normalize_space(line)
            if not s:
                continue
            low = s.lower()
            if any(re.search(p, low) for p in GARBAGE_PATTERNS):
                continue
            # skip obvious field labels
            if re.match(r"(?i)^(name|aka|aliases|dob|pob|nationality|address)\b", s):
                continue
            if likely.match(s) and len(s) >= 10:
                individuals.append({
                    "NAME": s,
                    "AKA": "",
                    "FOREIGN_SCRIPT": "",
                    "SEX": "",
                    "DOB": "",
                    "POB": "",
                    "NATIONALITY": "",
                    "OTHER_INFO": "Auto-extracted (fallback heuristic)",
                    "ADD": "",
                    "ADD_COUNTRY": "",
                    "TITLE": "",
                    "CITIZENSHIP": "",
                    "REMARK": "",
                })

    # De-duplicate by NAME
    seen = set()
    uniq = []
    for r in individuals:
        key = (r.get("NAME") or "").strip().lower()
        if key and key not in seen:
            seen.add(key)
            uniq.append(r)

    return uniq


# =========================
# XLSX Update
# =========================
def update_xlsx_with_individuals(xlsx_file, individuals: List[Dict[str, str]]) -> Tuple[bytes, int, int]:
    wb = load_workbook(xlsx_file)
    ws = wb.active

    # Detect headers from row 1 (if empty, use default columns)
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    headers = [h for h in headers if h] or COLUMNS

    existing_names = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = str(row[0]).strip().lower() if row and row[0] else ""
        if name:
            existing_names.add(name)

    added = 0
    skipped = 0

    for ind in individuals:
        name_key = (ind.get("NAME") or "").strip().lower()
        if not name_key or name_key in existing_names:
            skipped += 1
            continue

        next_row_num = ws.max_row + 1
        new_row = [ind.get(col, "") or "" for col in headers]

        for ci, value in enumerate(new_row, start=1):
            cell = ws.cell(row=next_row_num, column=ci, value=value if value else None)
            cell.font = Font(name="Calibri", size=11)

        existing_names.add(name_key)
        added += 1

    ws.freeze_panes = "A2"

    # Optional: preserve common widths (safe even if columns differ)
    col_widths = {
        "A": 49.36, "B": 39.82, "C": 14.82,
        "E": 14.18, "F": 11.82, "G": 8.54,
        "H": 33.45, "K": 7.18, "L": 6.82, "M": 13.18,
    }
    for col_letter, width in col_widths.items():
        try:
            ws.column_dimensions[col_letter].width = width
        except Exception:
            pass

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read(), added, skipped


# =========================
# UI
# =========================
st.markdown("## 🛡️ BH Sanctions List Updater (OFFLINE)")
st.caption("Upload Gazette PDF + BH‑TL‑INDIVIDUALS.xlsx → download updated XLSX (No API, no JSON).")
st.divider()

with st.sidebar:
    st.markdown("### OCR Settings (used only if needed)")
    dpi = st.slider("OCR DPI", min_value=150, max_value=400, value=250, step=25)
    ocr_lang = st.text_input("Tesseract language(s)", value="eng", help="Examples: eng, ara, eng+ara")

    st.markdown("---")
    st.markdown("### Debug")
    show_debug = st.toggle("Show diagnostics", value=True)

col1, col2 = st.columns(2)
with col1:
    pdf_file = st.file_uploader("📄 Upload Gazette PDF", type=["pdf"])
with col2:
    xlsx_file = st.file_uploader("📊 Upload BH‑TL‑INDIVIDUALS.xlsx", type=["xlsx", "xls"])

run = st.button("▶ Extract & Update XLSX", type="primary", disabled=not (pdf_file and xlsx_file))

if run and pdf_file and xlsx_file:
    pdf_bytes = pdf_file.read()

    text, method = extract_pdf_text_smart(pdf_bytes, dpi=dpi, lang=ocr_lang)

    if show_debug:
        with st.expander("🧪 Diagnostics", expanded=True):
            st.write(f"Extractor used: **{method}**")
            st.write(f"Extracted text length: **{len(text)}** characters")
            preview = (text[:2000] + (" …" if len(text) > 2000 else "")) if text else "[EMPTY]"
            st.code(preview, language="text")

            if method.startswith("OCR"):
                st.info(
                    "OCR was used because native text extraction looked empty/junk. "
                    "OCR requires Tesseract installed on the host."
                )

    if not text:
        st.error(
            "No extractable text was found even after OCR. "
            "If you are running this on a server, ensure Tesseract OCR is installed and accessible."
        )
        st.stop()

    individuals = parse_individual_blocks(text)

    st.info(f"Found **{len(individuals)}** unique candidate individual(s).")

    with st.expander("👀 Preview extracted names", expanded=False):
        if not individuals:
            st.write("No names extracted.")
        else:
            for i, r in enumerate(individuals[:50], start=1):
                st.write(f"{i}. {r.get('NAME','')}")
            if len(individuals) > 50:
                st.caption(f"Showing first 50 of {len(individuals)}")

    out_bytes, added, skipped = update_xlsx_with_individuals(xlsx_file, individuals)

    st.success(f"Done — Added: **{added}**, Skipped (duplicates/empty): **{skipped}**")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = xlsx_file.name.rsplit(".", 1)[0]
    out_name = f"{base}-UPDATED-{ts}.xlsx"

    st.download_button(
        "⬇️ Download Updated XLSX",
        data=out_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
