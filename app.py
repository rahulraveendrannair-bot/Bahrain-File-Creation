import io
import re
from datetime import datetime
from typing import List, Dict, Tuple

import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
from openpyxl import load_workbook
from openpyxl.styles import Font

# OCR is optional. If not installed / tesseract missing, app still runs with native extraction.
try:
    import pytesseract
    from PIL import Image
    OCR_AVAILABLE = True
except Exception:
    OCR_AVAILABLE = False


# =========================
# Streamlit page setup
# =========================
st.set_page_config(page_title="BH Sanctions List Updater", page_icon="🛡️", layout="centered")

COLUMNS = [
    "NAME", "AKA", "FOREIGN_SCRIPT", "SEX", "DOB", "POB",
    "NATIONALITY", "OTHER_INFO", "ADD", "ADD_COUNTRY",
    "TITLE", "CITIZENSHIP", "REMARK",
]

# Junk/navigation artifacts that can appear in extracted text
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

# Strong “field-like” patterns (may exist in some PDFs, not yours)
RE_NAME_LINE = re.compile(r"(?im)^\s*Name\s*:\s*(.+?)\s*$")

# Bahrain Gazette / UN style name parts show up as numbered tokens:
# e.g. "1: SAMI 2: JASIM 3: MUHAMMAD JAATA 4: AL-JABURI"
RE_NUMBERED_TOKEN = re.compile(r"(\d+)\s*:\s*([A-Z][A-Z \-']{1,80})")


# =========================
# Helpers
# =========================
def normalize_space(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def looks_like_garbage(text: str) -> bool:
    """
    Heuristic: treat extracted text as garbage if it's tiny or mostly nav text.
    """
    if not text or len(text.strip()) < 80:
        return True
    low = text.lower()
    hits = sum(1 for p in GARBAGE_PATTERNS if re.search(p, low))
    # if we see junk markers and text is small-ish, treat as garbage
    return hits >= 1 and len(text.strip()) < 1200


# =========================
# PDF Text extraction
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


def extract_text_ocr_tesseract(pdf_bytes: bytes, dpi: int = 250, lang: str = "eng") -> str:
    """
    OCR fallback for scanned/image-based PDFs.
    Requires pytesseract + PIL and the Tesseract binary installed on the host.
    """
    if not OCR_AVAILABLE:
        return ""

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


def extract_pdf_text_smart(pdf_bytes: bytes, dpi: int = 250, lang: str = "eng") -> Tuple[str, str]:
    """
    Returns (text, method_used)
      1) PyMuPDF native
      2) pdfplumber native
      3) OCR (if available) when text is missing / garbage
    """
    t1 = extract_text_pymupdf_native(pdf_bytes)
    if t1 and not looks_like_garbage(t1):
        return t1, "PyMuPDF (native)"

    t2 = extract_text_pdfplumber_native(pdf_bytes)
    if t2 and not looks_like_garbage(t2):
        return t2, "pdfplumber (native)"

    # OCR fallback
    if OCR_AVAILABLE:
        t3 = extract_text_ocr_tesseract(pdf_bytes, dpi=dpi, lang=lang)
        if t3:
            return t3, f"OCR (Tesseract) @ {dpi} DPI lang={lang}"

    # Last resort: return best we have (even if garbage) so diagnostics can show something
    if t1:
        return t1, "PyMuPDF (native, weak)"
    if t2:
        return t2, "pdfplumber (native, weak)"
    return "", "none"


# =========================
# Bahrain-specific parsing (fix)
# =========================
def extract_names_from_numbered_tokens(text: str) -> List[str]:
    """
    Extracts names from numbered tokens pattern:
      1: SAMI 2: JASIM 3: MUHAMMAD JAATA 4: AL-JABURI
    Handles tokens on the SAME line or across multiple lines.

    Strategy:
      - Scan text sequentially and detect sequences starting at 1:
          1:, 2:, 3: ... until break
      - Build a full name by joining token values in numeric order
    """
    tokens = RE_NUMBERED_TOKEN.findall(text)
    if not tokens:
        return []

    # Convert to (num, value)
    parsed = []
    for num_s, val in tokens:
        num = int(num_s)
        val = normalize_space(val)
        # Filter junk-ish token values
        low = val.lower()
        if any(re.search(p, low) for p in GARBAGE_PATTERNS):
            continue
        # Avoid collecting empty / super short fragments
        if len(val) < 2:
            continue
        parsed.append((num, val))

    names = []
    i = 0
    while i < len(parsed):
        num, val = parsed[i]
        # We only start a person when we see "1:"
        if num != 1:
            i += 1
            continue

        seq = {1: val}
        j = i + 1
        expected = 2
        while j < len(parsed):
            n2, v2 = parsed[j]
            if n2 == expected:
                seq[expected] = v2
                expected += 1
                j += 1
            elif n2 == 1:
                # a new sequence starts
                break
            else:
                # ignore out-of-order tokens until we find expected or a new 1:
                j += 1

        # Build name if we got at least 2 parts
        if len(seq) >= 2:
            full_name = " ".join(seq[k] for k in sorted(seq.keys()))
            full_name = normalize_space(full_name)
            # filter “Individuals click here” style lines if they somehow got through
            if not any(re.search(p, full_name.lower()) for p in GARBAGE_PATTERNS):
                names.append(full_name)

        i = j if j > i else i + 1

    # de-dup
    out = []
    seen = set()
    for n in names:
        k = n.lower()
        if k not in seen:
            seen.add(k)
            out.append(n)
    return out


def parse_individuals(text: str) -> List[Dict[str, str]]:
    """
    Build individuals list for XLSX from extracted PDF text.

    Priority:
      1) Bahrain numbered tokens (your PDF uses this) [1](https://wisetechglobal.sharepoint.com/sites/Content-as-Code/Shared%20Documents/SystemComponents/Home/eDocs-and-DocManager/Howto/How-to-call-Glow-API-with-DocToken-for-Shipamax-endpoints.aspx?web=1)
      2) If not found, fallback to Name: lines if present
      3) If still not found, return []
    """
    individuals: List[Dict[str, str]] = []

    # 1) Bahrain numbered token names
    numbered_names = extract_names_from_numbered_tokens(text)
    for nm in numbered_names:
        individuals.append({
            "NAME": nm,
            "AKA": "",
            "FOREIGN_SCRIPT": "",
            "SEX": "",
            "DOB": "",
            "POB": "",
            "NATIONALITY": "",
            "OTHER_INFO": "Auto-extracted from Bahrain Gazette (offline parser)",
            "ADD": "",
            "ADD_COUNTRY": "",
            "TITLE": "",
            "CITIZENSHIP": "",
            "REMARK": "",
        })

    if individuals:
        return individuals

    # 2) Fallback: Name: lines (some docs use this)
    for m in RE_NAME_LINE.finditer(text):
        nm = normalize_space(m.group(1))
        if nm and not any(re.search(p, nm.lower()) for p in GARBAGE_PATTERNS):
            individuals.append({
                "NAME": nm,
                "AKA": "",
                "FOREIGN_SCRIPT": "",
                "SEX": "",
                "DOB": "",
                "POB": "",
                "NATIONALITY": "",
                "OTHER_INFO": "Auto-extracted from PDF (Name: field)",
                "ADD": "",
                "ADD_COUNTRY": "",
                "TITLE": "",
                "CITIZENSHIP": "",
                "REMARK": "",
            })

    # de-dup
    seen = set()
    out = []
    for r in individuals:
        k = (r.get("NAME") or "").strip().lower()
        if k and k not in seen:
            seen.add(k)
            out.append(r)
    return out


# =========================
# XLSX update
# =========================
def update_xlsx_with_individuals(xlsx_file, individuals: List[Dict[str, str]]) -> Tuple[bytes, int, int]:
    wb = load_workbook(xlsx_file)
    ws = wb.active

    # Detect headers from row 1
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    headers = [h for h in headers if h] or COLUMNS

    existing = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        name = str(row[0]).strip().lower() if row and row[0] else ""
        if name:
            existing.add(name)

    added = skipped = 0

    for person in individuals:
        name = (person.get("NAME") or "").strip()
        key = name.lower()
        if not key or key in existing:
            skipped += 1
            continue

        next_row = ws.max_row + 1
        row_values = [person.get(col, "") or "" for col in headers]

        for ci, value in enumerate(row_values, start=1):
            cell = ws.cell(row=next_row, column=ci, value=value if value else None)
            cell.font = Font(name="Calibri", size=11)

        existing.add(key)
        added += 1

    ws.freeze_panes = "A2"

    # Preserve typical widths (safe even if columns differ)
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
    st.markdown("### Extraction settings")
    dpi = st.slider("OCR DPI (only if OCR used)", min_value=150, max_value=400, value=250, step=25)
    ocr_lang = st.text_input("Tesseract language(s)", value="eng", help="Examples: eng, ara, eng+ara")
    show_debug = st.toggle("Show diagnostics", value=True)

    if not OCR_AVAILABLE:
        st.warning("OCR libraries not available in this environment. Native extraction only.")

c1, c2 = st.columns(2)
with c1:
    pdf_file = st.file_uploader("📄 Upload Gazette PDF", type=["pdf"])
with c2:
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
                st.info("OCR was used because native extraction was missing/garbage.")

    if not text:
        st.error(
            "No extractable text found. If your PDF is scanned, OCR is needed and requires Tesseract installed."
        )
        st.stop()

    individuals = parse_individuals(text)

    st.info(f"Found **{len(individuals)}** candidate individual(s).")

    with st.expander("👀 Preview extracted names", expanded=True):
        if not individuals:
            st.write("No names extracted.")
        else:
            for i, r in enumerate(individuals[:100], start=1):
                st.write(f"{i}. {r.get('NAME', '')}")
            if len(individuals) > 100:
                st.caption(f"Showing first 100 of {len(individuals)}")

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
