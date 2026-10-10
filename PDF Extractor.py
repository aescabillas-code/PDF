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
import pytesseract
from PIL import Image
from io import BytesIO
from datetime import date

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

# Region-specific extraction labels only. No GL codes or GL mapping are stored in source.
REGION_LABELS = {
    "AU": {
        "net": ["NET AMOUNT", "NET AMOUNT (USD)", "NET AMOUNT(USD)"],
        "tax": ["TOTAL G.S.T", "TOTAL GST"],
    },
    "NA and LATAM": {
        "net": ["NET AMOUNT (USD)", "NET AMOUNT(USD)"],
        "tax": ["TOTAL REGION VAT/TAXES (USD)", "TOTAL REGION VAT/TAXES(USD)", "TOTAL REGION VAT/TAXES"],
    },
    "EMEA & APAC": {
        "net": ["NET AMOUNT (USD)", "NET AMOUNT(USD)"],
        "tax": ["TOTAL REGION VAT/TAXES (USD)", "TOTAL REGION VAT/TAXES(USD)", "TOTAL REGION VAT/TAXES"],
    },
    "UK": {
        "net": ["SUBTOTAL"],
        "tax": ["VAT"],
    },
}
REGION_ALIASES = {
    "NA/LATAM": "NA and LATAM",
    "NA & LATAM": "NA and LATAM",
    "EMEA APAC": "EMEA & APAC",
    "EMEA/APAC": "EMEA & APAC",
}
REGION_RULES = {region: {} for region in REGION_LABELS}
NET_LABEL = "Net amount / Subtotal"
VAT_LABEL = "Tax amount / VAT / GST"
MONEY_RE = re.compile(r"(?<![A-Z0-9])\$?\s*([\d,]+\.\d{2})(?!\d)")
EMAIL_RE = re.compile(r"[\\w.+-]+@[\\w.-]+\\.[A-Za-z]{2,}")
APPROVAL_RE = re.compile(r"Approved\\s+by\\s+(.+?)\\s+(\\d{2}[./-]\\d{2}[./-]\\d{4})", re.I)


def normalize_region(value: str) -> str:
    value = re.sub(r"\\s+", " ", (value or "").strip())
    return REGION_ALIASES.get(value.upper(), value)


def money_after(text: str, start: int = 0) -> Optional[float]:
    match = MONEY_RE.search(text, start)
    return float(match.group(1).replace(",", "")) if match else None


def get_approval_defaults():
    """Read default approver identity from Streamlit secrets without hardcoding it."""
    try:
        approval_secrets = st.secrets.get("approval", {})
        name = str(approval_secrets.get("approved_by", "")).strip()
        email = str(approval_secrets.get("email", "")).strip()
        return name, email
    except Exception:
        # Missing secrets should not prevent the app from running.
        return "", ""
APPROVAL_RE = re.compile(r"Approved\s+by\s+(.+?)\s+(\d{2}[./-]\d{2}[./-]\d{4})", re.I)


def normalize_region(value: str) -> str:
    value = re.sub(r"\s+", " ", (value or "").strip())
    return REGION_ALIASES.get(value.upper(), value)


def money_after(text: str, start: int = 0) -> Optional[float]:
    match = MONEY_RE.search(text, start)
    return float(match.group(1).replace(",", "")) if match else None


def _amount_near_label(text: str, labels: list[str]) -> Optional[float]:
    """Extract the amount associated with a label, avoiding other summary columns."""
    lines = [re.sub(r"\\s+", " ", line).strip() for line in (text or "").splitlines() if line.strip()]
    upper_lines = [line.upper() for line in lines]
    stop_labels = (
        "NET AMOUNT", "TOTAL G.S.T", "TOTAL GST", "TOTAL REGION VAT/TAXES",
        "SUBTOTAL", "VAT", "TOTAL AMOUNT", "INVOICE TOTAL", "AMOUNT DUE",
    )
    for index, line in enumerate(lines):
        upper = upper_lines[index]
        matched = next((label for label in sorted(labels, key=len, reverse=True) if label.upper() in upper), None)
        if not matched:
            continue
        label_match = re.search(re.escape(matched), line, re.I)
        if label_match:
            same_line = money_after(line, label_match.end())
            if same_line is not None:
                return same_line
        for next_index in range(index + 1, min(len(lines), index + 4)):
            candidate = upper_lines[next_index]
            if any(stop in candidate for stop in stop_labels):
                break
            values = list(MONEY_RE.finditer(lines[next_index]))
            if values:
                return float(values[0].group(1).replace(",", ""))
        for previous_index in range(index - 1, max(-1, index - 3), -1):
            if any(stop in upper_lines[previous_index] for stop in stop_labels):
                break
            values = list(MONEY_RE.finditer(lines[previous_index]))
            if values:
                return float(values[-1].group(1).replace(",", ""))
    # Handle flattened text by restricting the search to text after this label and
    # before the next financial label.
    compact = re.sub(r"\\s+", " ", text or "")
    label_pattern = r"NET\s*AMOUNT\s*\(?\s*USD\s*\)?|TOTAL\s+REGION\s+VAT\s*/\s*TAXES\s*\(?\s*USD\s*\)?|TOTAL\s+G\.?S\.?T\.?|TOTAL\s+GST|TOTAL\s+AMOUNT\s*\(?\s*USD\s*\)?|INVOICE\s+TOTAL|SUBTOTAL|\bVAT\b|AMOUNT\s+DUE"
    for label in sorted(labels, key=len, reverse=True):
        match = re.search(re.escape(label), compact, re.I)
        if match:
            tail = compact[match.end():]
            next_label = re.search(label_pattern, tail, re.I)
            segment = tail[:next_label.start()] if next_label else tail[:100]
            amount = money_after(segment)
            if amount is not None:
                return amount
    return None


def extract_labeled_gl_rows(text: str, region: str) -> tuple[list[dict], Optional[float]]:
    """Create stamp rows from region labels and GL codes configured in Streamlit secrets."""
    region = normalize_region(region)
    labels = REGION_LABELS.get(region)
    if not labels:
        return [], None

    net_amount = _amount_near_label(text, labels["net"])
    tax_amount = _amount_near_label(text, labels["tax"])

    # If the PDF text extractor puts table headers on one line and values on the
    # next, match the value position by the horizontal order of the header labels.
    lines = [re.sub(r"\\s+", " ", line).strip() for line in (text or "").splitlines() if line.strip()]
    for idx, line in enumerate(lines):
        upper = line.upper()
        net_pos = min((upper.find(x.upper()) for x in labels["net"] if upper.find(x.upper()) >= 0), default=-1)
        tax_pos = min((upper.find(x.upper()) for x in labels["tax"] if upper.find(x.upper()) >= 0), default=-1)
        if net_pos < 0 or tax_pos < 0:
            continue
        for next_line in lines[idx + 1:idx + 5]:
            values = list(MONEY_RE.finditer(next_line))
            if len(values) >= 2:
                net_value = float(values[0].group(1).replace(",", ""))
                tax_value = float(values[1].group(1).replace(",", ""))
                if net_amount is None:
                    net_amount = net_value
                if tax_amount is None:
                    tax_amount = tax_value
                break

    # Codes are retrieved only from secrets; no embedded GL-code fallback exists.
    codes = _region_gl_codes(region)
    rows = []
    if net_amount is not None and codes.get("net"):
        rows.append({"Juniper GL Code": codes["net"], "Amount": net_amount})
    if tax_amount is not None and codes.get("vat"):
        rows.append({"Juniper GL Code": codes["vat"], "Amount": tax_amount})
    return rows, None


def _region_gl_codes(region: str) -> dict:
    """Read GL codes exclusively from Streamlit secrets; never hardcode them."""
    region = normalize_region(region)
    try:
        secret_regions = st.secrets.get("regional_gl", {})
        for name, values in secret_regions.items():
            if normalize_region(name) == region and hasattr(values, "get"):
                return {
                    "net": str(values.get("net", "")).strip(),
                    "vat": str(values.get("vat", "")).strip(),
                }
    except Exception:
        pass
    return {"net": "", "vat": ""}


def classify_rows(region: str, rows: list[dict]) -> list[dict]:
    codes = _region_gl_codes(region)
    classified = []
    for row in rows:
        code = str(row.get("Juniper GL Code", "")).strip()
        if code and code == codes.get("net"):
            label = NET_LABEL
        elif code and code == codes.get("vat"):
            label = VAT_LABEL
        else:
            label = "Review — code not configured in secrets"
        classified.append({**row, "Classification": label})
    return classified


def validate_rows(region: str, rows: list[dict], total: Optional[float]) -> list[str]:
    warnings = []
    if normalize_region(region) not in REGION_LABELS:
        warnings.append("Select a configured region before stamping.")
    if not rows:
        warnings.append("No GL rows detected. Check region labels and regional_gl secrets.")
    codes = _region_gl_codes(region)
    if not codes.get("net") or not codes.get("vat"):
        warnings.append("Regional GL codes are missing from Streamlit secrets.")
    if any(row.get("Amount") is None for row in rows):
        warnings.append("At least one amount could not be paired confidently with its label.")
    if any(row["Classification"].startswith("Review") for row in rows):
        warnings.append("At least one GL code is not configured in secrets.")
    if total is None:
        warnings.append("Total amount was not detected.")
    elif rows and all(row.get("Amount") is not None for row in rows):
        calculated = round(sum(float(row["Amount"]) for row in rows), 2)
        if calculated != round(float(total), 2):
            warnings.append("Line amounts do not match the detected total. Review the values.")
    return warnings


def extract_from_page(text: str, region_hint: Optional[str] = None) -> dict:
    lines = [re.sub(r"\\s+", " ", line).strip() for line in (text or "").splitlines() if line.strip()]
    region = normalize_region(region_hint or "")
    invoice_match = re.search(r"INVOICE NUMBER\\s*([A-Z0-9-]+)", text or "", re.I)
    approver = approval_date = email = ""
    for line in lines:
        match = APPROVAL_RE.search(line)
        if match:
            approver, approval_date = match.group(1).strip(), match.group(2)
        email_match = EMAIL_RE.search(line)
        if email_match:
            email = email_match.group(0)
    rows, _ = extract_labeled_gl_rows(text, region)
    return {
        "region": region,
        "invoice_number": invoice_match.group(1) if invoice_match else "",
        "rows": rows,
        "total": None,
        "approver": approver,
        "approval_date": approval_date,
        "email": email,
    }


def _looks_like_summary(page_text: str) -> bool:
    upper_text = re.sub(r"\s+", " ", (page_text or "").upper())
    has_gl_block = "JUNIPER GL CODE" in upper_text and ("TOTAL AMOUNT" in upper_text or "APPROVED BY" in upper_text)
    labels = REGION_LABELS
    all_net = [term.upper() for item in labels.values() for term in item["net"]]
    all_tax = [term.upper() for item in labels.values() for term in item["tax"]]
    has_net = any(term in upper_text for term in all_net)
    has_tax = any(term in upper_text for term in all_tax)
    has_total = any(term in upper_text for term in ("TOTAL AMOUNT", "INVOICE TOTAL", "AMOUNT DUE"))
    has_region_pair = any(
        any(term.upper() in upper_text for term in item["net"])
        and any(term.upper() in upper_text for term in item["tax"])
        for item in REGION_LABELS.values()
    )
    return has_gl_block or has_region_pair or (has_net and has_tax and has_total)


def _looks_like_financial_summary(page_text: str) -> bool:
    """Recognize invoice totals pages even when no GL approval block exists."""
    upper_text = re.sub(r"\s+", " ", (page_text or "").upper())
    has_totals = any(label in upper_text for label in (
        "SUBTOTAL", "TOTAL G.S.T", "TOTAL GST", "TOTAL REGION VAT/TAXES",
        "TOTAL AMOUNT", "INVOICE TOTAL", "AMOUNT DUE", " VAT ", " AMOUNT "
    ))
    has_money = bool(MONEY_RE.search(page_text or ""))
    return has_totals and has_money


def _ocr_page(page) -> str:
    # Render in memory only; no invoice images or OCR text are written to disk.
    pix = page.get_pixmap(matrix=pymupdf.Matrix(2, 2), alpha=False)
    image = Image.open(BytesIO(pix.tobytes("png")))
    return pytesseract.image_to_string(image, config="--psm 6")


def find_candidate_pages(pdf_bytes: bytes) -> tuple[list[tuple[int, str]], bool]:
    """Find summary pages using embedded text first, then OCR as a fallback.

    Returns (candidate_pages, used_ocr). OCR runs across pages only when the
    normal text pass finds no summary, keeping text-based PDFs fast.
    """
    document = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        text_pages = []
        candidates = []
        financial_candidates = []
        for page_number, page in enumerate(document):
            page_text = page.get_text("text")
            text_pages.append((page_number, page_text))
            if _looks_like_summary(page_text):
                candidates.append((page_number, page_text))
            elif _looks_like_financial_summary(page_text):
                financial_candidates.append((page_number, page_text))
        if candidates:
            return candidates, False
        # Many vendor invoices contain a financial totals page but do not yet
        # contain the Juniper GL approval block. Allow manual GL entry on that page.
        if financial_candidates:
            return financial_candidates, False

        ocr_candidates = []
        ocr_financial_candidates = []
        ocr_available = True
        for page_number, page in enumerate(document):
            try:
                page_text = _ocr_page(page)
            except Exception:
                ocr_available = False
                # Continue to other pages; user-facing errors remain generic.
                continue
            if _looks_like_summary(page_text):
                ocr_candidates.append((page_number, page_text))
            elif _looks_like_financial_summary(page_text):
                ocr_financial_candidates.append((page_number, page_text))
        if ocr_candidates:
            return ocr_candidates, True
        if ocr_financial_candidates:
            return ocr_financial_candidates, True
        # Last-resort manual mode: expose a small set of final pages so the user
        # can choose where to place the stamp even if labels are unrecognized.
        nonempty_pages = [(n, t) for n, t in text_pages if t.strip()]
        if nonempty_pages:
            return nonempty_pages[-5:], not ocr_available
        return [], not ocr_available
    finally:
        document.close()


def classify_rows(region: str, rows: list[dict]) -> list[dict]:
    region_codes = _region_gl_codes(region)
    classified = []
    for row in rows:
        code = str(row["Juniper GL Code"])
        if code == str(region_codes.get("net", "")):
            label = NET_LABEL
        elif code == str(region_codes.get("vat", "")):
            label = VAT_LABEL
        else:
            label = "Review — region/code not mapped"
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
    """Place the stamp centered in a detected blank area without covering invoice content."""
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

        longest = max((len(line) for line in stamp_lines), default=30)
        width = min(max(330, longest * font_size * 0.53 + 20), page_rect.width - 56)
        height = len(stamp_lines) * line_gap + 12
        x = max(28, (page_rect.width - width) / 2)

        occupied = []
        for block in page.get_text("blocks"):
            x0, y0, x1, y1 = block[:4]
            if x1 > 0 and y1 > 0 and x0 < page_rect.width and y0 < page_rect.height:
                occupied.append(pymupdf.Rect(x0, y0, x1, y1))
        try:
            for drawing in page.get_drawings():
                drawing_rect = drawing.get("rect")
                if drawing_rect and drawing_rect.width > 0 and drawing_rect.height > 0:
                    occupied.append(pymupdf.Rect(drawing_rect))
        except Exception:
            pass

        center_y = (page_rect.height - height) / 2
        candidates = [center_y]
        step = max(12, line_gap)
        max_offset = int(max(0, page_rect.height - height - 36))
        for offset in range(step, max_offset + step, step):
            candidates.extend([center_y - offset, center_y + offset])

        def is_clear(y):
            candidate = pymupdf.Rect(x - 8, y - 6, x + width + 8, y + height + 6)
            if candidate.x0 < 20 or candidate.x1 > page_rect.width - 20:
                return False
            if candidate.y0 < 18 or candidate.y1 > page_rect.height - 18:
                return False
            for item in occupied:
                # Ignore thin page-border rules at the very edge.
                if item.width > page_rect.width * 0.92 and (item.y0 < 6 or item.y1 > page_rect.height - 6):
                    continue
                if candidate.intersects(item):
                    return False
            return True

        y = next((candidate for candidate in candidates if is_clear(candidate)), None)
        if y is None:
            raise ValueError("No clear space is available for the stamp on this page.")

        panel = pymupdf.Rect(x - 8, y - 6, x + width + 8, y + height + 6)
        page.draw_rect(panel, color=(0.82, 0.85, 0.87), fill=(1, 1, 1), width=0.7, overlay=True)
        for index, line in enumerate(stamp_lines):
            if line:
                page.insert_text(
                    (x + 4, y + index * line_gap + font_size),
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
    selected_region = st.selectbox("Invoice region", list(REGION_LABELS.keys()))
    stamp_placement = st.selectbox(
        "Stamp placement",
        ["Center in available blank space", "Bottom of summary page"],
    )
    st.caption("Region-specific validation rules are applied automatically.")

uploaded_pdf = st.file_uploader("Upload invoice PDF", type=["pdf"])

if uploaded_pdf is not None:
    pdf_content = uploaded_pdf.getvalue()  # in-memory; no local temp copy
    try:
        with st.spinner("Searching PDF text and using OCR if needed..."):
            candidate_pages, used_ocr = find_candidate_pages(pdf_content)
    except Exception:
        st.error("The PDF could not be opened. Check that it is a valid, readable PDF.")
        st.stop()

    if used_ocr:
        st.info("The PDF did not contain a detectable text summary in its embedded text. OCR was attempted; verify every extracted value before stamping.")
    if not candidate_pages:
        st.error("No readable page text was detected. You can still use the tool only after selecting a readable PDF or installing/configuring OCR for scanned pages.")
        st.stop()

    has_gl_summary = any(_looks_like_summary(page_text) for _, page_text in candidate_pages)
    if not has_gl_summary:
        st.warning("No pre-existing Juniper GL approval block was detected. Amounts are extracted from the selected region labels and GL codes are loaded from Streamlit secrets.")

    page_labels = [f"Page {page_number + 1}" for page_number, _ in candidate_pages]
    selected_page_label = st.selectbox(
        "Summary / approval page",
        page_labels,
        index=len(page_labels) - 1,
    )
    selected_position = page_labels.index(selected_page_label)
    page_index, page_text = candidate_pages[selected_position]
    extracted = extract_from_page(page_text, selected_region)

    # Preserve the reference script's extraction behavior. If the selected page
    # is only partially parsed, retry that same page with OCR before showing fields.
    if not extracted.get("rows") or extracted.get("total") is None:
        try:
            _doc = pymupdf.open(stream=pdf_content, filetype="pdf")
            try:
                _ocr_text = _ocr_page(_doc[page_index])
            finally:
                _doc.close()
            _ocr_extracted = extract_from_page(_ocr_text, selected_region)
            if len(_ocr_extracted.get("rows", [])) > len(extracted.get("rows", [])) or (
                _ocr_extracted.get("total") is not None and extracted.get("total") is None
            ):
                extracted = _ocr_extracted
                page_text = _ocr_text
                st.info("The selected page was re-read with OCR because the initial extraction was incomplete. Verify every value.")
        except Exception:
            pass

    st.subheader("Extracted information")
    st.caption(f"Selected summary page: {page_index + 1}")

    if not extracted.get("rows"):
        st.warning("No GL rows were extracted automatically. Confirm the selected region, label text, and regional_gl secrets.")

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

    secret_approver_name, secret_approver_email = get_approval_defaults()
    # Default approval date is the local date when extraction is generated.
    # Allow the user to edit it if the actual approval date differs.
    extraction_date = date.today().isoformat()

    first, second, third = st.columns(3)
    with first:
        # Always derive the displayed total from the current GL rows when amounts
        # are available, including after the user edits them.
        amount_values = (
            pd.to_numeric(edited_frame["Amount"], errors="coerce").dropna().tolist()
            if "Amount" in edited_frame.columns else []
        )
        if amount_values:
            default_total = round(sum(float(value) for value in amount_values), 2)
        elif extracted["total"] is not None:
            default_total = float(extracted["total"])
        else:
            default_total = 0.0
        invoice_total = st.number_input(
            "Total amount (sum of GL amounts)", min_value=0.0, value=default_total, step=0.01, format="%.2f"
        )
    with second:
        approver_default = extracted["approver"] or secret_approver_name
        approver_name = st.text_input("Approved by", value=approver_default, key="approved_by_input")
    with third:
        date_default = extraction_date
        approval_date = st.text_input("Approval date (defaults to extraction date)", value=date_default, key="approval_date_input")
    email_default = extracted["email"] or secret_approver_email
    approver_email = st.text_input("Approver email", value=email_default, key="approver_email_input")

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
                    file_name=f"{base_name}_with_Approval.pdf",
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
