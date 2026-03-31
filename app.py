import re
import io
import fitz  # PyMuPDF
import pdfplumber
import pytesseract
from PIL import Image

GARBAGE_PATTERNS = [
    r"click here",
    r"individuals\s+click",
    r"entities\s+click",
    r"back to",
    r"home",
    r"download",
    r"http[s]?://",
    r"www\.",
]

def looks_like_garbage(text: str) -> bool:
    if not text or len(text.strip()) < 50:
        return True
    t = text.lower()
    return any(re.search(p, t) for p in GARBAGE_PATTERNS)

def extract_text_native(pdf_bytes: bytes) -> tuple[str, str]:
    # 1) PyMuPDF native
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    parts = []
    for page in doc:
        t = page.get_text("text")
        if t:
            parts.append(t)
    text = "\n".join(parts).strip()
    if text:
        return text, "PyMuPDF(native)"

    # 2) pdfplumber native fallback
    parts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                parts.append(t)
    text = "\n".join(parts).strip()
    return text, "pdfplumber(native)" if text else ("", "none")

def extract_text_ocr(pdf_bytes: bytes, dpi: int = 250) -> tuple[str, str]:
    """
    OCR fallback:
    - Render each page to an image using PyMuPDF
    - OCR via pytesseract
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    parts = []
    zoom = dpi / 72  # PDF default DPI is 72
    mat = fitz.Matrix(zoom, zoom)

    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        page_text = pytesseract.image_to_string(img, lang="eng")  # add +ara if Arabic traineddata installed
        if page_text:
            parts.append(page_text)

    text = "\n".join(parts).strip()
    return text, "PyMuPDF(render)+Tesseract(OCR)"

def extract_pdf_text_smart(pdf_bytes: bytes) -> tuple[str, str]:
    text, method = extract_text_native(pdf_bytes)
    if looks_like_garbage(text):
        ocr_text, ocr_method = extract_text_ocr(pdf_bytes)
        if ocr_text and not looks_like_garbage(ocr_text):
            return ocr_text, ocr_method
    return text, method

def extract_names(text: str) -> list[str]:
    """
    Prefer "Name:" lines (Gazette / UN formats often include these).
    If not found, fallback to 'likely name' but filter garbage.
    """
    names = []
    lines = [re.sub(r"\s+", " ", ln).strip() for ln in text.splitlines() if ln.strip()]

    # 1) Strong pattern: Name:
    for ln in lines:
        m = re.match(r"(?i)^name\s*:\s*(.+)$", ln)
        if m:
            candidate = m.group(1).strip()
            if candidate and not any(re.search(p, candidate.lower()) for p in GARBAGE_PATTERNS):
                names.append(candidate)

    # 2) If none, fallback to likely name lines (Title Case / CAPS), but skip garbage
    if not names:
        likely = re.compile(r"^[A-Za-z][A-Za-z'.-]+(?:\s+[A-Za-z][A-Za-z'.-]+){1,6}$")
        for ln in lines:
            low = ln.lower()
            if any(re.search(p, low) for p in GARBAGE_PATTERNS):
                continue
            if likely.match(ln) and len(ln) >= 8:
                names.append(ln)

    # de-dup
    seen = set()
    out = []
    for n in names:
        k = n.lower()
        if k not in seen:
            seen.add(k)
            out.append(n)
    return out
