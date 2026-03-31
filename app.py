"""
BH Sanctions List Updater — Streamlit App
==========================================
Matches exact column format of BH-TL-INDIVIDUALS.xlsx
"""

import io
import json
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font

st.set_page_config(
    page_title="BH Sanctions List Updater",
    page_icon="🛡️",
    layout="wide",
)

COLUMNS = [
    "NAME", "AKA", "FOREIGN_SCRIPT", "SEX", "DOB", "POB",
    "NATIONALITY", "OTHER_INFO", "ADD", "ADD_COUNTRY",
    "TITLE", "CITIZENSHIP", "REMARK",
]

# ── Alias map: any key Claude returns → correct column ─────────────────────
ALIASES = {
    "name": "NAME", "full_name": "NAME", "full name": "NAME",
    "english_name": "NAME", "english name": "NAME",

    "aka": "AKA", "also_known_as": "AKA", "also known as": "AKA",
    "aliases": "AKA", "alias": "AKA", "other_names": "AKA", "other names": "AKA",

    "foreign_script": "FOREIGN_SCRIPT", "foreign script": "FOREIGN_SCRIPT",
    "arabic_name": "FOREIGN_SCRIPT", "arabic name": "FOREIGN_SCRIPT",
    "arabic": "FOREIGN_SCRIPT", "name_in_original_language": "FOREIGN_SCRIPT",
    "name in original language": "FOREIGN_SCRIPT", "original_script": "FOREIGN_SCRIPT",

    "sex": "SEX", "gender": "SEX",

    "dob": "DOB", "date_of_birth": "DOB", "date of birth": "DOB",
    "birth_date": "DOB", "birthdate": "DOB", "born": "DOB",

    "pob": "POB", "place_of_birth": "POB", "place of birth": "POB",
    "birth_place": "POB", "birthplace": "POB",

    "nationality": "NATIONALITY", "nationalities": "NATIONALITY",

    "other_info": "OTHER_INFO", "other info": "OTHER_INFO",
    "other_information": "OTHER_INFO", "other information": "OTHER_INFO",
    "additional_info": "OTHER_INFO", "additional_information": "OTHER_INFO",
    "details": "OTHER_INFO", "information": "OTHER_INFO",
    "notes": "OTHER_INFO", "additional_details": "OTHER_INFO",
    "listing_information": "OTHER_INFO", "listing information": "OTHER_INFO",

    "add": "ADD", "address": "ADD", "full_address": "ADD", "full address": "ADD",

    "add_country": "ADD_COUNTRY", "address_country": "ADD_COUNTRY",
    "address country": "ADD_COUNTRY", "country": "ADD_COUNTRY",
    "country_of_address": "ADD_COUNTRY",

    "title": "TITLE", "designation": "TITLE",

    "citizenship": "CITIZENSHIP", "citizenships": "CITIZENSHIP",

    "remark": "REMARK", "remarks": "REMARK", "note": "REMARK",
    "comments": "REMARK", "comment": "REMARK",
}


def normalize(raw: dict) -> dict:
    """Map any dict from Claude onto exact COLUMNS keys."""
    result = {col: "" for col in COLUMNS}
    for k, v in raw.items():
        key = k.strip().lower().replace("-", "_")
        mapped = ALIASES.get(key) or ALIASES.get(k.strip().lower())
        if mapped and v:
            val = str(v).strip()
            if result[mapped] and val:
                result[mapped] = result[mapped] + "; " + val
            elif val:
                result[mapped] = val
    return result


def parse_json(text: str) -> list[dict]:
    """Parse JSON from Claude, stripping markdown fences if present."""
    text = text.strip()
    if "```" in text:
        for part in text.split("```"):
            part = part.strip()
            if part.startswith("json"):
                part = part[4:].strip()
            if part.startswith("[") or part.startswith("{"):
                text = part
                break
    data = json.loads(text)
    if isinstance(data, dict):
        data = [data]
    if not isinstance(data, list):
        raise ValueError("JSON must be an array")
    return [normalize(row) for row in data if isinstance(row, dict)]


# ── CSS ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.block-container { padding-top: 1.5rem; }
.badge {
    background:#e8f4ff; color:#0066cc; border:1px solid #b3d4f5;
    border-radius:4px; padding:2px 8px; font-size:11px; font-weight:700;
    margin-left:8px; vertical-align:middle;
}
.metric-row { display:flex; gap:14px; margin:12px 0 16px; }
.mbox {
    background:#f0faf5; border:1px solid #b3e6cf;
    border-radius:8px; padding:14px 20px; text-align:center; flex:1;
}
.mnum { font-size:30px; font-weight:800; color:#00875a; }
.mlbl { font-size:11px; color:#666; margin-top:2px; }
.mbox.skip { background:#fffaf0; border-color:#f5d07a; }
.mbox.skip .mnum { color:#b07d00; }
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────
st.markdown(
    '## 🛡️ BH Sanctions List Updater <span class="badge">OFFLINE</span>',
    unsafe_allow_html=True,
)
st.caption("Upload Gazette PDF + BH-TL-INDIVIDUALS.xlsx + Claude-extracted JSON → Download updated XLSX")
st.divider()

# ── Extraction prompt ──────────────────────────────────────────────────────
with st.expander("📋 Prompt to use in Claude AI to get the JSON", expanded=False):
    st.markdown("Copy this prompt, then upload the Gazette PDF to [claude.ai](https://claude.ai) and send it:")
    st.code("""Extract all sanctioned individuals from this PDF as a JSON array.

For each person return an object with EXACTLY these keys:
- NAME: Full name in English (uppercase)
- AKA: All aliases separated by semicolons (include all name variants)
- FOREIGN_SCRIPT: Name in Arabic or other non-Latin script
- SEX: Male / Female or leave empty
- DOB: Date of birth (include all known dates separated by semicolons)
- POB: Place of birth
- NATIONALITY: Nationality
- OTHER_INFO: All other details — role, activities, listing date, passport numbers, ID numbers, physical description, alias document details, UN resolution references
- ADD: Full address
- ADD_COUNTRY: Country of address
- TITLE: Title or designation (Haji, Dr, Sheikh, etc.)
- CITIZENSHIP: Citizenship (if different from nationality)
- REMARK: Gazette issue number, date added to list, UN SC resolution, reference number

Return ONLY the JSON array. No markdown, no backticks, no explanation text.""",
    language="text")

    st.markdown("**Expected output format:**")
    st.code("""[
  {
    "NAME": "SAMI JASIM MUHAMMAD JAATA AL-JABURI",
    "AKA": "Mustafa Adnan al-Aziz; Mustafa Adnan al-Azeez; Sami al-Ajuz; Hajji Hamid",
    "FOREIGN_SCRIPT": "سامي جاسم محمد جعاطة الجبوري",
    "SEX": "Male",
    "DOB": "1 Jul. 1974",
    "POB": "Iraq",
    "NATIONALITY": "Iraqi",
    "OTHER_INFO": "QDi.437. Operated within ISIL (Da'esh) in Iraq and Syria. Listed in QDe.115. Head of Empowered Committee in ISIL. Financial affairs for ISIL. Terrorist operations against security forces. Mother's name: Aisha. Physical description: tall, brown skin, brown hair, black eyes. Identity docs: a) Syrian Arab Republic national ID 9080002892 (Mustafa Adnan al-Aziz, DOB 1 Jan 1973, POB Al-Bu Kamal, Syria); b) Turkish residence permit 4118 issued 15 Jan 2019 (Mustafa Adnan al-Azeez)",
    "ADD": "Iraq",
    "ADD_COUNTRY": "Iraq",
    "TITLE": "",
    "CITIZENSHIP": "Iraqi",
    "REMARK": "Gazette No. 3869, 29 March 2026. Added to list: 26 March 2026. UN SC Resolution 2734 (2024), Chapter VII UN Charter. Reference: DPPA/SCAD/SCSOB/2026/SCA/2026 (03)"
  }
]""", language="json")

st.divider()

# ── File uploaders ─────────────────────────────────────────────────────────
c1, c2, c3 = st.columns(3)
with c1:
    st.markdown("**📄 Gazette PDF**")
    pdf_file = st.file_uploader("PDF", type=["pdf"], label_visibility="collapsed", key="pdf")
    if pdf_file:
        st.success(f"✓ {pdf_file.name}")
with c2:
    st.markdown("**📊 BH-TL-INDIVIDUALS.xlsx**")
    xlsx_file = st.file_uploader("XLSX", type=["xlsx","xls"], label_visibility="collapsed", key="xlsx")
    if xlsx_file:
        st.success(f"✓ {xlsx_file.name}")
with c3:
    st.markdown("**{ } Individuals JSON**")
    json_file = st.file_uploader("JSON", type=["json","txt"], label_visibility="collapsed", key="json")
    if json_file:
        st.success(f"✓ {json_file.name}")

st.divider()

# ── JSON preview ───────────────────────────────────────────────────────────
if json_file:
    st.markdown("### 👁️ Preview — verify before updating")
    try:
        raw_text = json_file.read().decode("utf-8")
        json_file.seek(0)
        previews = parse_json(raw_text)

        if previews:
            for i, p in enumerate(previews):
                with st.expander(f"#{i+1} — {p.get('NAME','(no name)')}", expanded=True):
                    col_a, col_b = st.columns(2)
                    fields_left  = ["NAME","AKA","FOREIGN_SCRIPT","SEX","DOB","POB","NATIONALITY","TITLE","CITIZENSHIP"]
                    fields_right = ["OTHER_INFO","ADD","ADD_COUNTRY","REMARK"]
                    with col_a:
                        for f in fields_left:
                            v = p.get(f,"")
                            st.markdown(f"**{f}**")
                            st.text(v if v else "—")
                    with col_b:
                        for f in fields_right:
                            v = p.get(f,"")
                            st.markdown(f"**{f}**")
                            st.text(v if v else "—")

            # Check for empty columns
            df_check = pd.DataFrame(previews)
            empty_cols = [c for c in COLUMNS if not df_check[c].str.strip().any()]
            if empty_cols:
                st.warning(f"⚠️ These columns are empty for all records: **{', '.join(empty_cols)}**\n\n"
                           "If unexpected, re-extract from Claude using the prompt above.")
        else:
            st.warning("No individuals found in JSON.")
    except Exception as e:
        st.error(f"Could not parse JSON: {e}")
    st.divider()

# ── Run ────────────────────────────────────────────────────────────────────
all_ready = pdf_file and xlsx_file and json_file
run = st.button("▶  Update XLSX", disabled=not all_ready, type="primary")
if not all_ready:
    st.caption("⬆ Upload all three files above to enable.")

if run and all_ready:
    logs = []
    added = skipped = 0
    output_bytes = None
    error = None

    with st.spinner("Processing..."):

        # Parse JSON
        try:
            raw_text = json_file.read().decode("utf-8")
            individuals = parse_json(raw_text)
            logs.append(("ok", f"JSON parsed — {len(individuals)} individual(s) found"))
        except Exception as e:
            error = f"Invalid JSON: {e}"

        # Load XLSX
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

        # Append
        if not error:
            for person in individuals:
                name = person.get("NAME", "").strip()
                key = name.lower()
                if not key or key in existing:
                    logs.append(("warn", f"Skipped duplicate: {name or '(empty)'}"))
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

            # Preserve column widths
            for col_letter, width in {
                "A":49.36,"B":39.82,"C":14.82,
                "E":14.18,"F":11.82,"G":8.54,
                "H":33.45,"K":7.18,"L":6.82,"M":13.18,
            }.items():
                ws.column_dimensions[col_letter].width = width
            ws.freeze_panes = "A2"

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            output_bytes = buf.read()
            logs.append(("ok", f"Done — {added} added, {skipped} skipped"))

    # Log output
    with st.expander("📋 Processing log", expanded=True):
        for level, msg in logs:
            st.markdown(f"{'✅' if level == 'ok' else '⚠️'} {msg}")

    if error:
        st.error(f"**Error:** {error}")
    elif output_bytes:
        st.success("**Updated successfully!**")
        st.markdown(f"""
        <div class="metric-row">
            <div class="mbox"><div class="mnum">{added}</div><div class="mlbl">Individual(s) Added</div></div>
            <div class="mbox skip"><div class="mnum">{skipped}</div><div class="mlbl">Duplicate(s) Skipped</div></div>
        </div>""", unsafe_allow_html=True)

        base = xlsx_file.name.rsplit(".", 1)[0]
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            label="⬇️  Download Updated XLSX",
            data=output_bytes,
            file_name=f"{base}-UPDATED-{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
