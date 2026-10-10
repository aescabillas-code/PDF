"""
Juniper PDF Approval Stamp Tool — privacy-hardened Streamlit app.

Privacy controls:
- No invoice values or extracted personal/business data are written to logs.
- Uploaded PDFs are processed from in-memory bytes; no temporary invoice files are created.
- No real invoice, customer, approver, email, or GL values are hardcoded.
- Errors shown to the user are generic; detailed exception text is not displayed.
- The output PDF is returned as a download and is not saved by the app to a server path.

Run:
    py -m pip install -r requirements.txt
    py -m streamlit run app.py
"""
from __future__ import annotations

import re
import base64
import json
from typing import Optional

import streamlit as st
import pymupdf
import pandas as pd

st.set_page_config(page_title="PDF Approval Stamp Tool", page_icon="📄", layout="wide", initial_sidebar_state="expanded")

# Hide Streamlit's built-in menu, toolbar, and deploy controls in the app UI.
st.markdown("""
<style>
#MainMenu {visibility: hidden !important;}
header[data-testid="stHeader"] {visibility: hidden !important; height: 0 !important;}
[data-testid="stToolbar"] {visibility: hidden !important; height: 0 !important;}
[data-testid="stDecoration"] {visibility: hidden !important;}
[data-testid="stStatusWidget"] {visibility: hidden !important;}
</style>
""", unsafe_allow_html=True)

# Encoded configuration keeps organization-specific labels and GL suffix mappings
# out of plain sight in this source file. This is obfuscation, not encryption.
_RULES_B64 = "eyJyb2xlcyI6eyJJUExTIjp7InZhdF9zdWZmaXgiOiIyMSIsIm5ldF9zdWZmaXgiOiIxOSJ9LCJBVSI6eyJ2YXRfc3VmZml4IjoiMjMiLCJuZXRfc3VmZml4IjoiMTkifSwiTkEgYW5kIExBVEFNIjp7InZhdF9zdWZmaXgiOiIyMSIsIm5ldF9zdWZmaXgiOiIxOSJ9LCJVSyI6eyJ2YXRfc3VmZml4IjoiMjEiLCJuZXRfc3VmZml4IjoiMTkifSwiRU1FQSAmIEFQQUMiOnsidmF0X3N1ZmZpeCI6IjIxIiwibmV0X3N1ZmZpeCI6IjE5In19LCJhbGlhc2VzIjp7Ik5BL0xBVEFNIjoiTkEgYW5kIExBVEFNIiwiTkEgJiBMQVRBTSI6Ik5BIGFuZCBMQVRBTSIsIkVNRUEgQVBBQyI6IkVNRUEgJiBBUEFDIiwiRU1FQS9BUEFDIjoiRU1FQSAmIEFQQUMifSwibmV0X2xhYmVsIjoiTkNSIENFIG9uLXNpdGUgY29zdCAvIFN1YnRvdGFsIiwidmF0X2xhYmVsIjoiVkFUIC8gR1NUIn0="
try:
    _CONFIG = json.loads(base64.b64decode(_RULES_B64).decode("utf-8"))
except Exception:
    # Safe fallback retains functionality if the encoded configuration is damaged.
    _CONFIG = {"roles": {"IPLS": {"vat_suffix": "21", "net_suffix": "19"}, "AU": {"vat_suffix": "23", "net_suffix": "19"}, "NA and LATAM": {"vat_suffix": "21", "net_suffix": "19"}, "UK": {"vat_suffix": "21", "net_suffix": "19"}, "EMEA & APAC": {"vat_suffix": "21", "net_suffix": "19"}}, "aliases": {}}
REGION_RULES = _CONFIG["roles"]
REGION_ALIASES = {key.upper(): value for key, value in _CONFIG.get("aliases", {}).items()}
NET_LABEL = _CONFIG.get("net_label", "NCR CE on-site cost / Subtotal")
VAT_LABEL = _CONFIG.get("vat_label", "VAT / GST")
GL_RE = re.compile(r"(?<!\d)(\d{19})(?!\d)")
MONEY_RE = re.compile(r"\$?\s*([\d,]+\.\d{2})")
EMAIL_RE = re.compile(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}")
APPROVAL_RE = re.compile(r"Approved\s+by\s+(.+?)\s+(\d{2}[./-]\d{2}[./-]\d{4})", re.I)


def normalize_region(value: str) -> str:
    value = re.sub(r"\s+", " ", (value or "").strip())
    return REGION_ALIASES.get(value.upper(), value)


def money_after(text: str, start: int = 0) -> Optional[float]:
    match = MONEY_RE.search(text, start)
    return float(match.group(1).replace(",", "")) if match else None


def extract_from_page(text: str, region_hint: Optional[str] = None) -> dict:
    lines = [re.sub(r"\s+", " ", line).strip() for line in text.splitlines() if line.strip()]
    region = normalize_region(region_hint or "")

    if not region:
        for i, line in enumerate(lines):
            if "JUNIPER GL CODE" in line.upper():
                for candidate in reversed(lines[max(0, i - 5):i]):
                    for configured_region in REGION_RULES:
                        if candidate.lower() == configured_region.lower():
                            region = configured_region
                            break
                    if region:
                        break
                if region:
                    break

    invoice_match = re.search(r"INVOICE NUMBER\s*([A-Z0-9-]+)", text, re.I)
    invoice_number = invoice_match.group(1) if invoice_match else ""

    approver = approval_date = email = ""
    for line in lines:
        match = APPROVAL_RE.search(line)
        if match:
            approver, approval_date = match.group(1).strip(), match.group(2)
        email_match = EMAIL_RE.search(line)
        if email_match:
            email = email_match.group(0)

    rows = []
    active = False
    total = None
    for line in lines:
        upper = line.upper()
        if "JUNIPER GL CODE" in upper:
            active = True
            continue
        if active and "TOTAL AMOUNT" in upper:
            total = money_after(line)
            break
        if active and APPROVAL_RE.search(line):
            break
        if active:
            gl_match = GL_RE.search(line)
            if gl_match:
                rows.append({
                    "Juniper GL Code": gl_match.group(1),
                    "Amount": money_after(line, gl_match.end()),
                })

    if not rows:
        for line in lines:
            gl_match = GL_RE.search(line)
            if gl_match:
                rows.append({
                    "Juniper GL Code": gl_match.group(1),
                    "Amount": money_after(line, gl_match.end()),
                })

    if total is None:
        for line in lines:
            if "TOTAL AMOUNT" in line.upper():
                total = money_after(line)
                if total is not None:
                    break

    if not region:
        # Conservative inference only; the user can explicitly select the region.
        region = "AU" if any(row["Juniper GL Code"].endswith("23") for row in rows) else ""

    return {
        "region": region,
        "invoice_number": invoice_number,
        "rows": rows,
        "total": total,
        "approver": approver,
        "approval_date": approval_date,
        "email": email,
    }


def find_candidate_pages(pdf_bytes: bytes) -> list[tuple[int, str]]:
    document = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        candidates = []
        for page_number, page in enumerate(document):
            page_text = page.get_text("text")
            upper_text = page_text.upper()
            if "JUNIPER GL CODE" in upper_text and (
                "TOTAL AMOUNT" in upper_text or "APPROVED BY" in upper_text
            ):
                candidates.append((page_number, page_text))
        return candidates
    finally:
        document.close()


def classify_rows(region: str, rows: list[dict]) -> list[dict]:
    rules = REGION_RULES.get(region, {})
    classified = []
    for row in rows:
        code = str(row["Juniper GL Code"])
        if code.endswith("19"):
            label = NET_LABEL
        elif rules and code.endswith(rules["vat_suffix"]):
            label = VAT_LABEL
        else:
            label = "Review — region/code suffix not mapped"
        classified.append({**row, "Classification": label})
    return classified


def validate_rows(region: str, rows: list[dict], total: Optional[float]) -> list[str]:
    warnings = []
    if region not in REGION_RULES:
        warnings.append("Select a configured region before stamping.")
    if not rows:
        warnings.append("No GL rows detected.")
    if any(row["Amount"] is None for row in rows):
        warnings.append("At least one amount could not be paired confidently with its GL code.")
    if any(row["Classification"].startswith("Review") for row in rows):
        warnings.append("At least one GL suffix is not mapped for this region.")
    if total is None:
        warnings.append("Total amount was not detected.")
    elif rows and all(row["Amount"] is not None for row in rows):
        calculated = round(sum(float(row["Amount"]) for row in rows), 2)
        if calculated != round(float(total), 2):
            warnings.append("Line amounts do not match the detected total. Review the values.")
    return warnings


def stamp_pdf(
    pdf_bytes: bytes,
    page_index: int,
    rows: list[dict],
    total: float,
    approver: str,
    approval_date: str,
    email: str,
    placement: str,
) -> bytes:
    document = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        page = document[page_index]
        page_rect = page.rect
        font_size = 9
        line_gap = 12

        stamp_lines = ["JUNIPER GL CODE                                      AMOUNT"]
        for row in rows:
            stamp_lines.append(f'{row["Juniper GL Code"]}    ${float(row["Amount"]):,.2f}')
        stamp_lines.append(f"TOTAL AMOUNT    ${float(total):,.2f}")
        stamp_lines.append("")
        stamp_lines.append(f"Approved by {approver} {approval_date}".strip())
        if email.strip():
            stamp_lines.append(email.strip())

        anchor_y = None
        for block in page.get_text("blocks"):
            block_text = block[4].upper()
            if "TOTAL AMOUNT" in block_text or "INVOICE TOTAL" in block_text:
                anchor_y = max(anchor_y or 0, block[3])
        if anchor_y is None:
            anchor_y = page_rect.height * 0.62

        stamp_height = len(stamp_lines) * line_gap + 10
        x = 42
        y = page_rect.height - stamp_height - 24 if placement == "Bottom of summary page" else anchor_y + 10
        if y + stamp_height > page_rect.height - 12 or y < 10:
            raise ValueError("Insufficient safe space for stamp at selected location.")

        width = min(420, page_rect.width - x - 20)
        background = pymupdf.Rect(x - 4, y - 4, x + width, y + stamp_height)
        page.draw_rect(background, color=(1, 1, 1), fill=(1, 1, 1), overlay=True)
        for index, line in enumerate(stamp_lines):
            if line:
                page.insert_text(
                    (x, y + index * line_gap + font_size),
                    line,
                    fontname="helv",
                    fontsize=font_size,
                    color=(0, 0, 0),
                    overlay=True,
                )
        return document.tobytes(garbage=4, deflate=True)
    finally:
        document.close()


st.markdown("""
<style>
.block-container {padding-top: 1.5rem; max-width: 1400px;}
.hero {padding: 1.2rem 1.4rem; border-radius: 14px; background: linear-gradient(120deg,#006B4F,#008A69); color: white; margin-bottom: 1rem;}
.hero h1 {color:white; margin:0;}
.hero p {color:#e7fff6; margin:.35rem 0 0 0;}
</style>
<div class="hero">
<h1>PDF Approval Stamp Tool</h1>
<p>Extract Juniper GL details, validate approval totals, and stamp a separate copy of an invoice.</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("Processing rules")
    st.caption("Select the region for the invoice you are processing.")
    selected_region = st.selectbox("Invoice region", list(REGION_RULES.keys()))
    stamp_placement = st.selectbox(
        "Stamp placement",
        ["Below invoice total", "Bottom of summary page"],
    )
    st.caption("Region-specific validation rules are applied automatically.")

uploaded_pdf = st.file_uploader("Upload invoice PDF", type=["pdf"])

if uploaded_pdf is not None:
    pdf_content = uploaded_pdf.getvalue()  # in-memory; no local temp copy
    try:
        with st.spinner("Reading PDF text..."):
            candidate_pages = find_candidate_pages(pdf_content)
    except Exception:
        st.error("The PDF could not be opened. Check that it is a valid, readable PDF.")
        st.stop()

    if not candidate_pages:
        st.error("No selectable-text Juniper GL summary was found. Scanned PDFs require OCR.")
        st.stop()

    page_labels = [f"Page {page_number + 1}" for page_number, _ in candidate_pages]
    selected_page_label = st.selectbox(
        "Summary / approval page",
        page_labels,
        index=len(page_labels) - 1,
    )
    selected_position = page_labels.index(selected_page_label)
    page_index, page_text = candidate_pages[selected_position]
    extracted = extract_from_page(page_text, selected_region)

    st.subheader("Extracted information")
    st.caption(f"Selected summary page: {page_index + 1}")

    rows = extracted["rows"] or [{"Juniper GL Code": "", "Amount": None}]
    editable_frame = pd.DataFrame(classify_rows(selected_region, rows))
    edited_frame = st.data_editor(
        editable_frame,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Juniper GL Code": st.column_config.TextColumn("Juniper GL Code", required=True),
            "Amount": st.column_config.NumberColumn("Amount", format="$%.2f", min_value=0.0),
            "Classification": st.column_config.TextColumn("Classification", disabled=True),
        },
        key="gl_editor",
    )

    first, second, third = st.columns(3)
    with first:
        default_total = float(extracted["total"]) if extracted["total"] is not None else 0.0
        invoice_total = st.number_input(
            "Total amount", min_value=0.0, value=default_total, step=0.01, format="%.2f"
        )
    with second:
        approver_name = st.text_input("Approved by", value=extracted["approver"])
    with third:
        approval_date = st.text_input("Approval date", value=extracted["approval_date"])
    approver_email = st.text_input("Approver email", value=extracted["email"])

    final_rows = []
    for _, row in edited_frame.iterrows():
        code = str(row.get("Juniper GL Code", "")).strip()
        amount = row.get("Amount")
        if code and amount is not None and str(amount).strip().lower() != "nan":
            final_rows.append({"Juniper GL Code": code, "Amount": float(amount)})

    classified_rows = classify_rows(selected_region, final_rows)
    warnings = validate_rows(selected_region, classified_rows, invoice_total)
    for warning in warnings:
        st.warning(warning)
    if not warnings:
        st.success("Line amounts match the total. Review the stamp preview before processing.")

    st.subheader("Stamp preview")
    preview_lines = ["JUNIPER GL CODE                         AMOUNT"]
    for row in final_rows:
        preview_lines.append(f'{row["Juniper GL Code"]}    ${row["Amount"]:,.2f}')
    preview_lines.append(f"TOTAL AMOUNT    ${invoice_total:,.2f}")
    preview_lines += ["", f"Approved by {approver_name} {approval_date}".strip(), approver_email]
    st.code("\n".join(preview_lines), language="text")

    if st.button("Create stamped PDF", type="primary", use_container_width=True):
        if not final_rows:
            st.error("Add at least one GL code and amount before stamping.")
        elif not approver_name.strip() or not approval_date.strip():
            st.error("Approver name and approval date are required.")
        elif warnings:
            st.error("Resolve the validation warnings before stamping.")
        else:
            try:
                result_pdf = stamp_pdf(
                    pdf_content,
                    page_index,
                    final_rows,
                    invoice_total,
                    approver_name,
                    approval_date,
                    approver_email,
                    stamp_placement,
                )
                original_name = uploaded_pdf.name
                base_name = original_name[:-4] if original_name.lower().endswith(".pdf") else original_name
                st.success("Stamped PDF is ready. Review the downloaded copy before using it.")
                st.download_button(
                    "Download stamped PDF",
                    data=result_pdf,
                    file_name=f"{base_name}_STAMPED.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            except ValueError:
                st.error("There is not enough safe space for the stamp at the selected location.")
            except Exception:
                st.error("Stamping failed. Check that the PDF is valid and that the selected page is correct.")
else:
    st.info("Upload a supported regional invoice PDF to begin.")
    st.markdown("""
    **Workflow**
    1. Select the invoice region.
    2. Upload the PDF and select its summary page.
    3. Review or correct the extracted GL codes and amounts.
    4. Confirm the total, approver, date, and email.
    5. Create and download the stamped PDF.

    Uploaded content is processed in memory by this app; it does not intentionally write invoice data to local temporary files or logs.
    """)
