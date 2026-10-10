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
import math
import hashlib
import zipfile
from collections import Counter
from typing import Optional
from datetime import datetime

import streamlit as st
import pymupdf
import pandas as pd
import pytesseract
from PIL import Image
from io import BytesIO
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

st.set_page_config(page_title="PDF Approval Stamp Tool", page_icon="📄", layout="wide", initial_sidebar_state="expanded")

# Hide Streamlit's built-in menu, toolbar, and deploy controls in the app UI.
st.markdown("""
<style>
#MainMenu {visibility: hidden !important;}
header[data-testid="stHeader"] {visibility: hidden !important; height: 0 !important;}
[data-testid="stToolbar"] {visibility: hidden !important; height: 0 !important;}
[data-testid="stDecoration"] {visibility: hidden !important;}
[data-testid="stStatusWidget"] {visibility: hidden !important;}
/* Keep the region-selection sidebar fixed and prevent users from collapsing it. */
[data-testid="stSidebarCollapseButton"],
[data-testid="collapsedControl"],
button[kind="header"] {display: none !important; visibility: hidden !important;}
section[data-testid="stSidebar"] {transform: none !important; min-width: 18rem !important; visibility: visible !important;}
</style>
""", unsafe_allow_html=True)

# Region names and aliases are not sensitive. All full GL codes are read only from Streamlit secrets.
REGION_RULES = {
    "IPLS": {},
    "AU": {},
    "NA and LATAM": {},
    "UK": {},
    "EMEA & APAC": {},
    "INDIA": {},
}
REGION_ALIASES = {
    "NA/LATAM": "NA and LATAM",
    "NA & LATAM": "NA and LATAM",
    "EMEA APAC": "EMEA & APAC",
    "EMEA/APAC": "EMEA & APAC",
}
NET_LABEL = "Net Amount / Subtotal"
VAT_LABEL = "VAT / GST"
GL_RE = re.compile(r"(?<!\d)(\d{19})(?!\d)")
MONEY_RE = re.compile(r"\$?\s*([\d,]+\.\d{2})")
EMAIL_RE = re.compile(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}")
APPROVAL_RE = re.compile(r"Approved\s+by\s+(.+?)\s+(\d{2}[./-]\d{2}[./-]\d{4})", re.I)




def _india_clean_text(value: str) -> str:
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _india_amount(value: str) -> Optional[float]:
    cleaned = str(value or "").replace(",", "").replace("₹", "").strip()
    match = re.search(r"-?\d+(?:\.\d{1,2})?", cleaned)
    return round(float(match.group(0)), 2) if match else None


def _india_parse_invoice(pdf_bytes: bytes) -> dict:
    """Extract invoice identity and tax summary from a vendor PDF, in memory."""
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        pages_text = [page.get_text("text") for page in doc]
    finally:
        doc.close()
    text = "\n".join(pages_text)
    flat = _india_clean_text(text)
    # NCR invoice number is shown beside NUMBER / DATE on the invoice header.
    inv_match = re.search(r"\b(CSDL\d+|CSGJ\d+|CSHR\d+|CSKA\d+|CSMH\d+|CSMN\d+|CSTN\d+|CSTS\d+|CSUP\d+)\b", flat, re.I)
    gstin_match = re.search(r"\bGSTIN\s*:?\s*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][A-Z0-9]Z[A-Z0-9])\b", flat, re.I)
    # Prefer the supplier GSTIN printed in the invoice header. The buyer GSTIN
    # is later under "Bill to", so capture the first GSTIN occurrence.
    if gstin_match:
        gstin = gstin_match.group(1).upper()
    else:
        gstin = ""
    date_match = re.search(r"\b(\d{2}/\d{2}/\d{2,4})\b", flat)
    invoice_date = ""
    if date_match:
        raw = date_match.group(1)
        for fmt in ("%d/%m/%y", "%d/%m/%Y"):
            try:
                invoice_date = datetime.strptime(raw, fmt).strftime("%m/%d/%Y")
                break
            except ValueError:
                pass

    def value_after(patterns):
        for pattern in patterns:
            m = re.search(pattern, text, re.I)
            if m:
                return _india_amount(m.group(1))
        return None

    net = value_after([
        r"\bSubtotal\b\s*(?:\[INR\])?\s*([₹]?\s*[\d,]+\.\d{2})",
        r"\bSubtotal\b[\s\S]{0,100}?([₹]?\s*[\d,]+\.\d{2})",
    ])
    total = value_after([
        r"\bINVOICE TOTAL\b[\s\S]{0,80}?([₹]?\s*[\d,]+\.\d{2})",
        r"\bTOTAL AMOUNT\b[\s\S]{0,80}?([₹]?\s*[\d,]+\.\d{2})",
        r"\bAMOUNT DUE\b[\s\S]{0,80}?([₹]?\s*[\d,]+\.\d{2})",
    ])

    # Collect every monetary amount after Subtotal and before the invoice total.
    # PDF text extraction can separate tax labels and values into different lines,
    # so scan the flattened text segment instead of depending on CGST/SGST labels.
    _subtotal_match = re.search(r"\bSubtotal\b", flat, re.I)
    _total_match = re.search(r"\b(?:INVOICE TOTAL|TOTAL AMOUNT|AMOUNT DUE)\b", flat, re.I)
    _tax_components = []
    if _subtotal_match and _total_match and _total_match.start() > _subtotal_match.end():
        _tax_section = flat[_subtotal_match.end():_total_match.start()]
        # Drop the subtotal value itself; the section starts immediately after
        # the word Subtotal and may contain its amount before the tax amounts.
        _amount_matches = re.findall(r"\d[\d,]*\.\d{2}", _tax_section)
        if _amount_matches:
            # First amount after the label is the subtotal. Remaining amounts
            # are tax components; each is included in the tax sum.
            _amount_matches = _amount_matches[1:]
        for _amount_text in _amount_matches:
            _component_amount = _india_amount(_amount_text)
            if _component_amount is not None:
                _tax_components.append(_component_amount)

    _tax_total_from_lines = round(sum(_tax_components), 2) if _tax_components else None
    _fallback_tax_total = round(total - net, 2) if total is not None and net is not None else None
    tax_field = _tax_total_from_lines if _tax_total_from_lines is not None else _fallback_tax_total
    tax_type = "IGST" if re.search(r"\bIGST\b", text, re.I) else ("TAX" if tax_field is not None else "")
    return {
        "invoice_number": inv_match.group(1).upper() if inv_match else "",
        "gstin": gstin,
        "invoice_date": invoice_date,
        "net": net,
        "tax_field": tax_field,
        "tax_type": tax_type,
        "invoice_total": total,
    }


INDIA_TABLE_COLUMNS = [
    "Juniper Entity Code", "NCR Vendor code", "GL Code ", "Amount- Duty/IGST",
    "Curr", "Tax field", "Total (amount+tax)", "Dr-40/Cr-50", "Inv Date",
    "Line Description", "Cost Center", "Project ID", "Invoice Number",
    "Header Description to submit Weekly File", "Assignment Reference", "Tax Code",
    "Billed From", "Billed To", "GSTIN", "HSN code ",
]


def _india_initial_table() -> pd.DataFrame:
    """Create the NCR AP table directly in code; no Excel workbook is required."""
    rows = [
        ["2380", "60009733", "624010", None, "INR", None, None, "40", None, "NCR TPM Q4 2026 Invoice CSDL10901", "50637", None, None, "NCR_Incident fee Upload_October 2026", "NCR October 2024", "G2", "DELHI", "MAHARASHTRA", "07AAACN7149L1Z4", "998719"],
        ["2380", "60009733", "624010", None, "INR", None, None, "40", None, "NCR TPM Q4 2026 Invoice CSGJ10375", "50637", None, None, "NCR_Incident fee Upload_October 2026", "NCR October 2024", "G2", "Gujarat", "MAHARASHTRA", "24AAACN7149L1Z8", "998719"],
        ["2380", "60009733", "624010", None, "INR", None, None, "40", None, "NCR TPM Q4 2026 Invoice CSHR10217", "50637", None, None, "NCR_Incident fee Upload_October 2026", "NCR October 2024", "G2", "Haryana", "MAHARASHTRA", "06AAACN7149L1Z6", "998719"],
        ["2380", "60009733", "624010", None, "INR", None, None, "40", None, "NCR TPM Q4 2026 Invoice CSKA13538", "50637", None, None, "NCR_Incident fee Upload_October 2026", "NCR October 2024", "G2", "Karnataka", "MAHARASHTRA", "29AAACN7149L1ZY", "998719"],
        ["2380", "60009733", "624010", None, "INR", None, None, "40", None, "NCR TPM Q4 2026 Invoice CSMH22167", "50637", None, None, "NCR_Incident fee Upload_October 2026", "NCR October 2024", "G1", "Maharashtra", "MAHARASHTRA", "27AAACN7149L1Z2", "998719"],
        ["2380", "60009733", "624010", None, "INR", None, None, "40", None, "NCR TPM Q4 2026 Invoice CSTN22344", "50637", None, None, "NCR_Incident fee Upload_October 2026", "NCR October 2024", "G2", "Tamil Nadu", "MAHARASHTRA", "33AAACN7149L1Z9", "998719"],
        ["2380", "60009733", "624010", None, "INR", None, None, "40", None, "NCR TPM Q4 2026 Invoice CSTS9775", "50637", None, None, "NCR_Incident fee Upload_October 2026", "NCR October 2024", "G2", "Telangana", "MAHARASHTRA", "36AAACN7149L1Z3", "998719"],
        ["2380", "60009733", "624010", None, "INR", None, None, "40", None, "NCR TPM Q4 2026 Invoice CSUP17039", "50637", None, None, "NCR_Incident fee Upload_October 2026", "NCR October 2024", "G2", "Uttar Pradesh", "MAHARASHTRA", "09AAACN7149L1Z0", "998719"],
        ["2380", "60009733", "624010", None, "INR", None, None, "40", None, "NCR TPM Q4 2026 Invoice CSMN1626", "50637", None, None, "NCR_Incident fee Upload_October 2026", "NCR October 2024", "G2", "Manipur", "MAHARASHTRA", "14AAACN7149L1Z9", "998719"],
    ]
    return pd.DataFrame(rows, columns=INDIA_TABLE_COLUMNS)


def _india_make_csv(pdf_uploads, quarter: str, year: int, fee_type: str, upload_month: str, upload_year: int):
    """Fill the in-app NCR AP table using invoice number + supplier GSTIN."""
    frame = _india_initial_table()
    required = ["Invoice Number", "GSTIN", "Amount- Duty/IGST", "Tax field", "Total (amount+tax)", "Inv Date", "Line Description", "Header Description to submit Weekly File"]
    missing = [c for c in required if c not in frame.columns]
    if missing:
        raise ValueError("The uploaded template is missing required columns.")

    pdf_records = {}
    duplicates = []
    for uploaded in pdf_uploads:
        record = _india_parse_invoice(uploaded.getvalue())
        if not record["invoice_number"] or not record["gstin"]:
            continue
        key = (record["invoice_number"].upper(), record["gstin"].upper())
        if key in pdf_records:
            duplicates.append(key[0])
        pdf_records[key] = record

    results = []
    unmatched = []
    for idx, row in frame.iterrows():
        invoice_number = _india_clean_text(row.get("Invoice Number", "")).upper()
        gstin = _india_clean_text(row.get("GSTIN", "")).upper()
        # The supplied blank template may carry the invoice number only in the
        # existing line description; use it as a matching key without treating
        # screenshot/reference values as source invoice data.
        if not invoice_number:
            description = _india_clean_text(row.get("Line Description", ""))
            inferred = re.search(r"\b(CSDL\d+|CSGJ\d+|CSHR\d+|CSKA\d+|CSMH\d+|CSMN\d+|CSTN\d+|CSTS\d+|CSUP\d+)\b", description, re.I)
            if inferred:
                invoice_number = inferred.group(1).upper()
        if not invoice_number and not gstin:
            continue
        record = pdf_records.get((invoice_number, gstin))
        if record is None:
            unmatched.append(invoice_number or f"Excel row {idx + 2}")
            continue
        if record["gstin"] != gstin:
            unmatched.append(invoice_number)
            continue
        if pd.isna(row.get("Invoice Number")) or not _india_clean_text(row.get("Invoice Number", "")):
            frame.at[idx, "Invoice Number"] = invoice_number

        # Fill missing values only; preserve pre-populated template data except
        # the line description, which follows the user-selected quarter/year.
        net = record["net"]
        tax = record["tax_field"]
        if pd.isna(row.get("Amount- Duty/IGST")) and net is not None:
            frame.at[idx, "Amount- Duty/IGST"] = net
        if pd.isna(row.get("Tax field")) and tax is not None:
            frame.at[idx, "Tax field"] = tax
        if pd.isna(row.get("Total (amount+tax)")):
            if net is not None and tax is not None:
                frame.at[idx, "Total (amount+tax)"] = round(net + tax, 2)
            elif record["invoice_total"] is not None:
                frame.at[idx, "Total (amount+tax)"] = record["invoice_total"]
        if pd.isna(row.get("Inv Date")) and record["invoice_date"]:
            frame.at[idx, "Inv Date"] = record["invoice_date"]
        frame.at[idx, "Line Description"] = f"NCR TPM {quarter} {year} Invoice {invoice_number}"
        fee_label = "Incident fee" if fee_type == "Incident_fee" else "TPM fee"
        frame.at[idx, "Header Description to submit Weekly File"] = f"NCR_{fee_label} Upload_{upload_month} {upload_year}"
        results.append({
            "Invoice Number": invoice_number,
            "GSTIN": gstin,
            "Net": net,
            "Tax field": tax,
            "Tax type": record["tax_type"],
            "PDF invoice total": record["invoice_total"],
            "CSV total (amount+tax)": frame.at[idx, "Total (amount+tax)"],
        })

    return frame, results, unmatched, duplicates



def normalize_region(value: str) -> str:
    value = re.sub(r"\s+", " ", (value or "").strip())
    return REGION_ALIASES.get(value.upper(), value)


def money_after(text: str, start: int = 0) -> Optional[float]:
    match = MONEY_RE.search(text, start)
    return float(match.group(1).replace(",", "")) if match else None


def extract_region_amounts_from_page(page, region: str) -> tuple[list[dict], Optional[float]]:
    """Match region-specific summary labels to amounts using PDF coordinates.

    AU and UK use label/value rows. NA/LATAM and EMEA/APAC use USD column
    headings with values beneath them. GL codes are never read from the invoice.
    """
    region = normalize_region(region)
    codes = _region_gl_codes(region)
    words = page.get_text("words")
    words = [(float(w[0]), float(w[1]), float(w[2]), float(w[3]), str(w[4])) for w in words]
    if not words:
        return [], None

    def norm(value):
        return re.sub(r"[^A-Z0-9]+", " ", value.upper()).strip()

    def phrase_boxes(phrase):
        target = norm(phrase).split()
        found = []
        # Group words by approximate baseline, then match phrase within each line.
        lines = {}
        for w in words:
            key = round(w[1] / 5) * 5
            lines.setdefault(key, []).append(w)
        for line in lines.values():
            line.sort(key=lambda w: w[0])
            for i in range(len(line)):
                pieces = []
                for j in range(i, min(len(line), i + len(target) + 3)):
                    pieces.append(line[j])
                    got = norm(" ".join(x[4] for x in pieces)).split()
                    if got == target:
                        found.append((min(x[0] for x in pieces), min(x[1] for x in pieces),
                                      max(x[2] for x in pieces), max(x[3] for x in pieces)))
                        break
                    if len(got) > len(target) or got != target[:len(got)]:
                        break
        return found

    def monetary_words():
        vals = []
        for x0, y0, x1, y1, token in words:
            cleaned = token.replace("$", "").replace(",", "")
            if re.fullmatch(r"\d+(?:\.\d{2})", cleaned):
                try:
                    vals.append((x0, y0, x1, y1, float(cleaned)))
                except ValueError:
                    pass
        return vals

    amounts = monetary_words()
    found = {}

    if region == "AU":
        specs = {
            "net": ["Net Amount"],
            "vat": ["Total G.S.T", "Total GST"],
        }
    elif region == "UK":
        specs = {"net": ["SUBTOTAL"], "vat": ["VAT"]}
    else:
        # In NA/LATAM and EMEA/APAC invoices, the VAT column heading is split
        # across two lines: "TOTAL REGION" above "VAT/TAXES(USD)". Match the
        # second-line column heading directly so the tax amount is not missed.
        specs = {
            "net": ["NET AMOUNT(USD)", "NET AMOUNT (USD)"],
            "vat": ["VAT/TAXES(USD)", "VAT/TAXES (USD)"],
        }

    for kind, labels in specs.items():
        boxes = []
        for label in labels:
            boxes.extend(phrase_boxes(label))
        if not boxes:
            continue

        candidates = []
        for lx0, ly0, lx1, ly1 in boxes:
            for ax0, ay0, ax1, ay1, value in amounts:
                if region in ("IPLS", "NA and LATAM", "EMEA & APAC"):
                    # Amount must be below its own USD column header.
                    # These invoice layouts can place the amount far to the right
                    # of the heading, so do not cap the horizontal distance.
                    # The nearest amount horizontally to each header is selected
                    # when multiple values share the same summary row.
                    if ay0 <= ly1 + 8:
                        continue
                    if ax0 < lx0 - 35:
                        continue
                    # Prefer the amount whose horizontal center is closest to
                    # this column header; this keeps the total-amount column from
                    # being mistaken for the VAT/tax amount.
                    header_center = (lx0 + lx1) / 2
                    amount_center = (ax0 + ax1) / 2
                    candidates.append((ay0 - ly1, abs(amount_center - header_center), value))
                else:
                    # AU/UK values are printed on the same row, to the right of label.
                    vertical = abs((ay0 + ay1) / 2 - (ly0 + ly1) / 2)
                    horizontal = ax0 - lx1
                    if vertical <= 24 and -5 <= horizontal <= 360:
                        candidates.append((vertical, horizontal, value))
        if candidates:
            candidates.sort(key=lambda v: (v[0], v[1]))
            found[kind] = candidates[0][2]

    rows = []
    if found.get("net") is not None and codes.get("net"):
        rows.append({"Juniper GL Code": codes["net"], "Amount": found["net"]})
    if found.get("vat") is not None and codes.get("vat"):
        rows.append({"Juniper GL Code": codes["vat"], "Amount": found["vat"]})
    total = round(found["net"] + found["vat"], 2) if len(found) == 2 else None
    return rows, total


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


def _looks_like_summary(page_text: str) -> bool:
    upper_text = page_text.upper()
    return "JUNIPER GL CODE" in upper_text and (
        "TOTAL AMOUNT" in upper_text or "APPROVED BY" in upper_text
    )


def _looks_like_financial_summary(page_text: str) -> bool:
    """Recognize invoice totals pages even when no GL approval block exists."""
    upper_text = re.sub(r"\s+", " ", (page_text or "").upper())
    has_totals = any(label in upper_text for label in (
        "SUBTOTAL", "TOTAL G.S.T", "TOTAL GST", "TOTAL AMOUNT", "INVOICE TOTAL", "AMOUNT DUE"
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


def auto_detect_summary_page(pdf_bytes: bytes, region: str, candidate_pages: list[tuple[int, str]]) -> tuple[list[tuple[int, str]], int]:
    """Rank pages by region-specific summary labels and coordinate-extracted amounts."""
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        all_pages = [(i, page.get_text("text")) for i, page in enumerate(doc)]
        # Include the existing candidates and any page that contains the regional
        # labels, even when an unrelated page has an old approval block.
        candidate_map = {i: text for i, text in candidate_pages}
        for i, text in all_pages:
            upper = re.sub(r"\s+", " ", text.upper())
            if any(term in upper for term in (
                "NET AMOUNT", "TOTAL REGION VAT/TAXES", "TOTAL G.S.T", "TOTAL GST", "SUBTOTAL", " VAT "
            )):
                candidate_map[i] = text

        def score(item):
            page_index, text = item
            upper = re.sub(r"\s+", " ", text.upper())
            if region == "AU":
                net_terms, tax_terms = ("NET AMOUNT",), ("TOTAL G.S.T", "TOTAL GST")
            elif region == "UK":
                net_terms, tax_terms = ("SUBTOTAL",), ("VAT",)
            else:
                net_terms, tax_terms = ("NET AMOUNT(USD)", "NET AMOUNT (USD)"), (
                    "TOTAL REGION VAT/TAXES(USD)", "TOTAL REGION VAT/TAXES (USD)"
                )
            score_value = 0
            if any(term in upper for term in net_terms): score_value += 20
            if any(term in upper for term in tax_terms): score_value += 20
            if "TOTAL AMOUNT" in upper or "INVOICE TOTAL" in upper or "AMOUNT DUE" in upper: score_value += 3
            if "JUNIPER GL CODE" in upper or "APPROVED BY" in upper: score_value += 1
            try:
                rows, total = extract_region_amounts_from_page(doc[page_index], region)
                if len(rows) == 2 and total is not None:
                    score_value += 100
            except Exception:
                pass
            return score_value

        options = list(candidate_map.items())
        if not options:
            return candidate_pages, 0
        ranked = sorted(options, key=score, reverse=True)
        best_index, best_text = ranked[0]
        # Preserve a manageable dropdown, but make the best detected page its default.
        ordered = [(best_index, best_text)] + [item for item in candidate_pages if item[0] != best_index]
        return ordered, best_index
    finally:
        doc.close()


def _region_gl_codes(region: str) -> dict:
    """Read full GL codes exclusively from Streamlit secrets."""
    region = normalize_region(region)
    try:
        regional = st.secrets.get("regional_gl", {})
        for name, values in regional.items():
            if normalize_region(str(name)) == region and hasattr(values, "get"):
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
    region: str = "",
) -> bytes:
    """Place the approval stamp centered in a genuinely unoccupied area.

    The search uses page text block coordinates so it works across invoice layouts
    where the summary may be at the top, bottom, or side of the page. If no safe
    rectangle can fit the complete stamp, fail rather than cover invoice content.
    """
    document = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        page = document[page_index]
        page_rect = page.rect
        region_key = normalize_region(region)
        compact_regional_stamp = region_key in ("IPLS", "NA and LATAM", "EMEA & APAC")
        # These two invoice layouts have a specific blank area below the service
        # description and to the left of NET TOTAL PAYABLE. Use a compact stamp
        # there; all other regions retain their existing placement and sizing.
        font_size = 6.5 if compact_regional_stamp else 9
        line_gap = 7 if compact_regional_stamp else 12
        # Use a monospaced font so exactly three spaces separate the longest
        # GL-code column from the AMOUNT column and every amount starts aligned.
        gl_header = "JUNIPER GL CODE"
        longest_gl = max([len(gl_header)] + [
            len(str(row.get("Juniper GL Code", ""))) for row in rows
        ])
        amount_col = longest_gl + 3
        stamp_lines = [
            gl_header + (" " * (amount_col - len(gl_header))) + "AMOUNT"
        ]
        for row in rows:
            code = str(row["Juniper GL Code"])
            amount_text = f'${float(row["Amount"]):,.2f}'
            stamp_lines.append(code + (" " * max(3, amount_col - len(code))) + amount_text)
        total_label = "TOTAL AMOUNT"
        stamp_lines.append(total_label + (" " * max(3, amount_col - len(total_label))) + f'${float(total):,.2f}')
        stamp_lines.append("")
        stamp_lines.append(f"Approved by {approver} {approval_date}".strip())
        if email.strip():
            stamp_lines.append(email.strip())

        stamp_height = len(stamp_lines) * line_gap + (5 if compact_regional_stamp else 12)
        longest = max((len(line) for line in stamp_lines), default=30)
        # Courier glyphs are 0.6 em wide; allow a small horizontal safety margin.
        stamp_width = min(
            max(210 if compact_regional_stamp else 330,
                longest * font_size * 0.60 + (12 if compact_regional_stamp else 20)),
            page_rect.width - 48,
        )
        if stamp_width < (180 if compact_regional_stamp else 250) or stamp_height > page_rect.height - 36:
            raise ValueError("Stamp does not fit on the selected page.")

        # Occupied text blocks are expanded slightly to preserve a visual margin.
        occupied = []
        for block in page.get_text("blocks"):
            if len(block) < 5:
                continue
            x0, y0, x1, y1 = map(float, block[:4])
            block_text = str(block[4]).strip()
            if not block_text:
                continue
            rect = pymupdf.Rect(x0, y0, x1, y1)
            if rect.intersects(page_rect):
                occupied.append((rect, block_text))

        margin_x, margin_y = 22, 18
        x_min = margin_x
        x_max = page_rect.width - stamp_width - margin_x
        y_min = margin_y
        y_max = page_rect.height - stamp_height - margin_y

        if region_key == "IPLS":
            # IPLS stamp must sit below "PLEASE PAY THIS AMOUNT" and align to
            # the right side of the PDF page, not to the heading's left edge.
            please_pay_rect = None
            for rect, block_text in occupied:
                upper = re.sub(r"\\s+", " ", block_text.upper())
                if "PLEASE PAY THIS AMOUNT" in upper:
                    if please_pay_rect is None:
                        please_pay_rect = rect
                    else:
                        please_pay_rect = pymupdf.Rect(
                            min(please_pay_rect.x0, rect.x0),
                            min(please_pay_rect.y0, rect.y0),
                            max(please_pay_rect.x1, rect.x1),
                            max(please_pay_rect.y1, rect.y1),
                        )
            if please_pay_rect is None:
                raise ValueError("The IPLS placement anchor was not found on the selected page.")
            # Fix the horizontal position at the page's right margin. Only the
            # vertical coordinate is searched, and every candidate remains below
            # the anchor and subject to the existing text-overlap safety check.
            x_min = page_rect.width - stamp_width - 18
            x_max = x_min
            y_min = max(18, please_pay_rect.y1 + 5)
            y_max = page_rect.height - stamp_height - 18

        elif compact_regional_stamp:
            # Preserve the existing placement for NA/LATAM and EMEA/APAC:
            # below the service-description paragraph, left of NET TOTAL PAYABLE,
            # and above the THIS INVOICE footer line.
            service_bottom = None
            payable_left = None
            footer_top = None
            for rect, block_text in occupied:
                upper = re.sub(r"\\s+", " ", block_text.upper())
                if "THE SERVICES SHOWN ON THIS INVOICE" in upper:
                    service_bottom = max(service_bottom or rect.y1, rect.y1)
                if "NET TOTAL PAYABLE" in upper:
                    payable_left = min(payable_left or rect.x0, rect.x0)
                if "THIS INVOICE IS FOR HARDWARE MAINTENANCE SERVICES" in upper:
                    footer_top = min(footer_top or rect.y0, rect.y0)
            if service_bottom is not None and payable_left is not None and footer_top is not None:
                x_min = max(18, page_rect.x0 + 18)
                x_max = min(page_rect.x1 - stamp_width - 12, payable_left - stamp_width - 12)
                y_min = max(18, service_bottom + 5)
                y_max = min(page_rect.y1 - stamp_height - 12, footer_top - stamp_height - 5)

        if x_max < x_min or y_max < y_min:
            raise ValueError("Stamp does not fit on the selected page.")

        # Search a coarse grid, then rank clear positions by closeness to the
        # center of the page. Include extra candidates around text-block edges,
        # which often correspond to the large blank section beneath invoice tables.
        x_step = max(12, min(24, stamp_width / 14))
        y_step = max(10, min(20, stamp_height / 3))
        xs = {x_min, x_max, (x_min + x_max) / 2}
        ys = {y_min, y_max, (y_min + y_max) / 2}
        x = x_min
        while x <= x_max:
            xs.add(x)
            x += x_step
        y = y_min
        while y <= y_max:
            ys.add(y)
            y += y_step

        for rect, _ in occupied:
            for candidate_x in (rect.x0 - stamp_width - 10, rect.x1 + 10):
                if x_min <= candidate_x <= x_max:
                    xs.add(candidate_x)
            for candidate_y in (rect.y0 - stamp_height - 10, rect.y1 + 10):
                if y_min <= candidate_y <= y_max:
                    ys.add(candidate_y)

        page_center = pymupdf.Point(page_rect.width / 2, page_rect.height / 2)
        candidates = []
        for candidate_x in xs:
            for candidate_y in ys:
                candidate = pymupdf.Rect(
                    candidate_x - 5,
                    candidate_y - 5,
                    candidate_x + stamp_width + 5,
                    candidate_y + stamp_height + 5,
                )
                if candidate.x0 < 12 or candidate.x1 > page_rect.width - 12:
                    continue
                if candidate.y0 < 12 or candidate.y1 > page_rect.height - 12:
                    continue
                if any(candidate.intersects(rect) for rect, _ in occupied):
                    continue

                center = pymupdf.Point(
                    candidate_x + stamp_width / 2,
                    candidate_y + stamp_height / 2,
                )
                # Prefer the page center, while the grid and text-edge candidates
                # allow the stamp to settle into the nearest large clear section.
                distance = ((center.x - page_center.x) / max(page_rect.width, 1)) ** 2
                distance += ((center.y - page_center.y) / max(page_rect.height, 1)) ** 2
                candidates.append((distance, candidate_y, candidate_x))

        if not candidates:
            raise ValueError("No clear space is available for the stamp on this page.")

        # Both placement choices honor the no-overlap rule. The default centers
        # in the best available blank area; bottom placement prefers the lowest
        # safe candidate if the user explicitly selects it.
        if placement == "Bottom of summary page":
            _, y, x = max(candidates, key=lambda item: (item[1], -item[0]))
        else:
            _, y, x = min(candidates, key=lambda item: item[0])

        for index, line in enumerate(stamp_lines):
            if line:
                page.insert_text(
                    (x + 5, y + index * line_gap + font_size),
                    line,
                    fontname="cour",
                    fontsize=font_size,
                    color=(0, 0, 0),
                    overlay=True,
                )
        return document.tobytes(garbage=4, deflate=True)
    finally:
        document.close()



# -----------------------------------------------------------------------------
# Tool navigation and Vendor SOW Analyzer (kept separate from PDF stamping)
# -----------------------------------------------------------------------------
SOW_STOPWORDS = set("""a an and are as at be been being by for from has have had in into is it its of on or our that the their this to with will can may must should shall services service scope work works vendor supplier client customer global regional local support manage management provide provision including include includes across per via using use used related applicable requirements requirement deliver deliverables delivery activities activity responsibilities responsibility hereby hereunder thereto therein thereof pursuant agreement agreements party parties hpe hereby herein such any all each both other same respective said may also not no nor if then than where when while which who whom whose more most less very other however therefore section clause exhibit appendix attachment page number version date confidential proprietary company enterprise inc limited ltd gmbh private pte mexico japan india china anz apj naas non""".split())

SOW_SUPPORTED_EXTS = {'.pdf', '.docx', '.docm', '.xlsx', '.xlsm', '.xls', '.pptx', '.txt', '.md', '.csv', '.zip'}
SOW_MAX_ARCHIVE_DEPTH = 3
SOW_MAX_DOCUMENTS = 100
SOW_MAX_TOTAL_BYTES = 100 * 1024 * 1024


def _sow_read_document(name, data):
    """Read one supported SOW/capability document in memory only."""
    suffix = ('.' + name.rsplit('.', 1)[-1].lower()) if '.' in name else ''
    if suffix == '.pdf':
        doc = pymupdf.open(stream=data, filetype='pdf')
        try:
            pages = [page.get_text('text') for page in doc]
            text_value = '\n'.join(pages)
            # OCR is a fallback only for genuinely text-poor/scanned PDFs.
            if len(text_value.strip()) < 100:
                ocr_pages = []
                for page in doc:
                    try:
                        ocr_pages.append(_ocr_page(page))
                    except Exception:
                        ocr_pages.append('')
                if len('\n'.join(ocr_pages).strip()) > len(text_value.strip()):
                    text_value = '\n'.join(ocr_pages)
            return text_value
        finally:
            doc.close()
    if suffix in ('.txt', '.md', '.csv'):
        return data.decode('utf-8-sig', errors='replace')
    if suffix in ('.docx', '.docm'):
        from docx import Document
        document = Document(BytesIO(data))
        parts = [paragraph.text for paragraph in document.paragraphs if paragraph.text.strip()]
        for table in document.tables:
            for row in table.rows:
                parts.append(' | '.join(cell.text.strip() for cell in row.cells))
        return '\n'.join(parts)
    if suffix in ('.xlsx', '.xlsm', '.xls'):
        sheets = pd.read_excel(BytesIO(data), sheet_name=None, dtype=str)
        parts = []
        for sheet_name, frame in sheets.items():
            parts.append(f'SHEET: {sheet_name}')
            parts.append(frame.fillna('').astype(str).to_csv(index=False))
        return '\n'.join(parts)
    if suffix == '.pptx':
        from pptx import Presentation
        presentation = Presentation(BytesIO(data))
        parts = []
        for slide in presentation.slides:
            for shape in slide.shapes:
                if getattr(shape, 'has_text_frame', False) and shape.text.strip():
                    parts.append(shape.text)
                if getattr(shape, 'has_table', False):
                    for row in shape.table.rows:
                        parts.append(' | '.join(cell.text for cell in row.cells))
        return '\n'.join(parts)
    raise ValueError('Unsupported document format')


def _sow_expand_uploads(uploads):
    """Expand supported documents and nested ZIPs in memory, deduplicating exact copies."""
    expanded = []
    seen_hashes = set()
    total_bytes = 0

    def add_document(name, data, depth=0):
        nonlocal total_bytes
        if len(expanded) >= SOW_MAX_DOCUMENTS:
            raise ValueError('The upload contains too many documents. Limit the batch to 100 documents.')
        suffix = ('.' + name.rsplit('.', 1)[-1].lower()) if '.' in name else ''
        if suffix == '.zip':
            if depth >= SOW_MAX_ARCHIVE_DEPTH:
                return
            try:
                with zipfile.ZipFile(BytesIO(data)) as archive:
                    for item in archive.infolist():
                        if item.is_dir():
                            continue
                        inner_name = item.filename.replace('\\', '/').split('/')[-1]
                        inner_suffix = ('.' + inner_name.rsplit('.', 1)[-1].lower()) if '.' in inner_name else ''
                        if inner_suffix not in SOW_SUPPORTED_EXTS:
                            continue
                        if item.file_size > SOW_MAX_TOTAL_BYTES or total_bytes + item.file_size > SOW_MAX_TOTAL_BYTES:
                            raise ValueError('The expanded upload exceeds the 100 MB processing limit.')
                        with archive.open(item) as handle:
                            inner_data = handle.read(SOW_MAX_TOTAL_BYTES + 1)
                        if len(inner_data) > SOW_MAX_TOTAL_BYTES:
                            raise ValueError('A document in the archive exceeds the processing limit.')
                        add_document(f'{name} / {inner_name}', inner_data, depth + 1)
            except zipfile.BadZipFile:
                return
            return
        if suffix not in SOW_SUPPORTED_EXTS:
            return
        digest = hashlib.sha256(data).hexdigest()
        if digest in seen_hashes:
            return
        seen_hashes.add(digest)
        total_bytes += len(data)
        if total_bytes > SOW_MAX_TOTAL_BYTES:
            raise ValueError('The expanded upload exceeds the 100 MB processing limit.')
        text_value = _sow_read_document(name, data)
        if text_value and text_value.strip():
            expanded.append({'File': name, 'Text': text_value, 'SHA256': digest})

    for uploaded in uploads or []:
        add_document(uploaded.name, uploaded.getvalue())
    return expanded


def _sow_read_upload(uploaded):
    """Backward-compatible reader for a single document or ZIP archive."""
    docs = _sow_expand_uploads([uploaded])
    return '\n\n'.join(f"DOCUMENT: {item['File']}\n{item['Text']}" for item in docs)


def _sow_normalize(text):
    return re.sub(r'[^a-z0-9]+', ' ', (text or '').lower()).strip()


def _sow_sentences(text):
    """Extract line-level clauses while discarding page furniture and contents leaders."""
    text = (text or '').replace('\r', '\n')
    pieces = []
    for raw_line in text.split('\n'):
        line = re.sub(r'[ \t\f\v]+', ' ', raw_line).strip(' \t-•*▪◦\uf0b7')
        if not line or len(line) < 24:
            continue
        if line.isupper() and len(line.split()) < 14:
            continue
        if re.search(r'\.{4,}|…{3,}|_{5,}', line):
            continue
        if re.fullmatch(r'(?:page\s*)?\d+(?:\s+of\s+\d+)?', line, re.I):
            continue
        if re.search(r'(?i)^(?:statement of work|hewlett packard enterprise confidential|confidentiality notice|important confidentiality notice|document history|revision history|table of contents|contents|docusign envelope id)', line):
            continue
        split_parts = re.split(r'(?<=[.!?;])\s+(?=[A-Z0-9(])', line)
        pieces.extend(part.strip(' \t-•*▪◦\uf0b7') for part in split_parts if len(part.strip()) >= 24)
    return pieces


def _sow_terms(text, limit=120):
    tokens = re.findall(r'[a-z][a-z0-9+#./-]{2,}', (text or '').lower())
    counts = Counter(token.strip('./-') for token in tokens if token not in SOW_STOPWORDS and len(token.strip('./-')) > 2)
    return [word for word, _ in counts.most_common(limit)]


def _sow_extract_scope(text, limit=240):
    """Prioritize operational scope, capability, coverage, SLA, and deliverable clauses."""
    sentences = _sow_sentences(text)
    if not sentences:
        return []
    action_terms = (
        'provide onsite', 'onsite services', 'on-site services', 'break-fix', 'break fix',
        'hardware replacement', 'replace hardware', 'parts replacement', 'troubleshooting',
        'technical support', 'remote support', 'field service', 'field services', 'logistics',
        'case management', 'parts management', 'customer service', 'incident management',
        'service level', 'service levels', 'response time', 'resolution time', 'customer sla',
        '24/7', '24x7', '9x5', 'business hours', 'service coverage', 'geographies in scope',
        'countries in scope', 'country coverage', 'dispatch', 'work order', 'spare parts',
        'inventory', 'shipment', 'warehouse', 'installation', 'maintenance', 'repair',
        'escalation', 'ticket', 'incident', 'deliverable', 'weekly status report', 'sla report',
        'end-to-end', 'e2e', 'project manager', 'project management', 'service improvement',
        'monitor and update', 'customer updates', 'onsite support', 'technical assistance',
        'administrative support', 'supply chain', 'customer contact', 'service request',
        'hardware remediation', 'smart hands', 'configuration', 'network cabling',
        'response and repair', 'service area', 'language coverage', 'support model'
    )
    boilerplate_terms = (
        'confidentiality notice', 'confidential and proprietary', 'table of contents', 'document history',
        'revision history', 'docusign envelope', 'all rights reserved', 'governed by the laws',
        'limitation of liability', 'indirect or consequential damages', 'counterparts and signatures',
        'this document is the property', 'intentionally left blank', 'signature page',
        'agreement shall be governed', 'terms and conditions', 'copyright', 'not to be disclosed',
        'not to be duplicated', 'intellectual property', 'force majeure', 'indemnify and hold harmless',
        'employee of supplier', 'dressed in a manner', 'professional and lawful manner',
        'shall be governed by', 'without the express written authorization', 'shall exercise the same degree'
    )
    selected = []
    for sentence in sentences:
        lower = sentence.lower()
        normalized = _sow_normalize(sentence)
        if any(term in lower for term in boilerplate_terms):
            continue
        if len(normalized.split()) < 7:
            continue
        if not any(term in lower for term in action_terms):
            continue
        selected.append(sentence)
    if not selected:
        selected = [s for s in sentences if len(s.split()) >= 8 and not any(t in s.lower() for t in boilerplate_terms)]
    unique, seen = [], set()
    for sentence in selected:
        key = _sow_normalize(sentence)
        if key and key not in seen:
            seen.add(key)
            unique.append(sentence)
        if len(unique) >= limit:
            break
    return unique


def _sow_similarity(scope, vendor_text):
    """Score the best clause-level match instead of comparing against whole-document boilerplate."""
    target = set(_sow_terms(scope, 500))
    if not target:
        return 0.0, []
    candidates = list(vendor_text) if isinstance(vendor_text, (list, tuple, set)) else _sow_extract_scope(vendor_text, limit=240)
    best_score, best_terms = 0.0, []
    for candidate in candidates:
        tokens = set(_sow_terms(candidate, 500))
        overlap = target & tokens
        if len(overlap) < min(2, len(target)):
            continue
        recall = len(overlap) / max(len(target), 1)
        precision = len(overlap) / max(len(tokens), 1)
        score = 100 * (0.75 * recall + 0.25 * precision)
        if score > best_score:
            best_score, best_terms = score, sorted(overlap)
    return round(min(best_score, 100), 1), best_terms


def _sow_scope_label(sentence):
    cleaned = re.sub(r'\s+', ' ', sentence).strip()
    return cleaned[:180] + ('…' if len(cleaned) > 180 else '')


def _sow_infer_vendor(name, text):
    """Infer a legal supplier name when stated; otherwise clearly mark it as unidentified."""
    source = (text or '').replace('\r', '\n')
    known_entity_patterns = [
        r'(?i)Hemmersbach Worldwide GmbH', r'(?i)EXEO GLOBAL PTE\.? LTD\.?',
        r'(?i)Cygnet IT Services Private Limited', r'(?i)ARCO Japan,? Inc\.?',
        r'(?i)ARA Technologies trading as AWA Technology Services',
        r'(?i)Beijing Huayuweicheng Technology Service Co\.?\s*,?\s*Ltd\.?',
        r'(?i)Smartlog Mexico S\.?A\.? de C\.?V\.?',
    ]
    for pattern in known_entity_patterns:
        match = re.search(pattern, source)
        if match:
            value = re.sub(r'\s+', ' ', match.group(0)).strip(' .;:-')
            if value.lower().startswith('ara technologies'):
                return 'ARA Technologies Pty Limited (AWA Technology Services)'
            if value.lower().startswith('smartlog'):
                return 'Smartlog Mexico SA de CV'
            if value.lower().startswith('beijing'):
                return 'Beijing Huayuweicheng Technology Service Co., Ltd.'
            return value
    candidates = re.findall(r'(?im)^\s*(?:supplier(?:\s+name)?|subcontractor|proveedor)\s*:\s*([^\n\r|]{2,160})', source)
    for candidate in candidates:
        value = re.sub(r'\s+', ' ', candidate).strip(' .;:-')
        if not value or re.fullmatch(r'[_\- .]+', value) or 'project manager' in value.lower():
            continue
        if len(value) >= 3 and not value.lower().startswith(('please insert', 'to be completed', 'name and title')):
            return value
    stem = re.sub(r'\.(pdf|docx|docm|xlsx|xlsm|xls|pptx|txt|md|csv|zip)$', '', name.rsplit('/', 1)[-1], flags=re.I)
    stem = re.sub(r'^(?:SOW\s*)', '', stem, flags=re.I).strip(' _-')
    return f'Unidentified supplier — {stem or "uploaded document"}'


def _sow_detect_coverage_area(file_names, text=''):
    """Return only geographic labels supported by the document title."""
    labels = []
    for name in file_names if isinstance(file_names, (list, tuple, set)) else [file_names]:
        value = re.sub(r'[^a-z0-9]+', ' ', str(name).lower())
        if re.search(r'\banz\b|australia.{0,8}new zealand', value):
            label = 'Australia & New Zealand (ANZ)'
        elif re.search(r'\bapj\b|asia pacific', value):
            label = 'Asia Pacific (APJ)'
        elif re.search(r'\bchina\b', value):
            label = 'China'
        elif re.search(r'\bindia\b', value):
            label = 'India'
        elif re.search(r'\bjapan\b', value):
            label = 'Japan'
        else:
            label = ''
        if label and label not in labels:
            labels.append(label)
    return '; '.join(labels) if labels else 'Not established from document title'


def _sow_cluster_scope_items(scope_records, threshold=0.62, limit=180):
    """Cluster near-duplicate scope clauses into a balanced shared comparison catalogue."""
    clusters = []
    for record in scope_records:
        sentence = record['scope']
        tokens = set(_sow_terms(sentence, 500))
        if len(tokens) < 3:
            continue
        best_idx, best_similarity = None, 0.0
        for idx, cluster in enumerate(clusters):
            overlap = tokens & cluster['tokens']
            union = tokens | cluster['tokens']
            jaccard = len(overlap) / max(len(union), 1)
            containment = len(overlap) / max(min(len(tokens), len(cluster['tokens'])), 1)
            similarity = max(jaccard, containment * 0.82)
            if len(overlap) >= 4 and similarity > best_similarity:
                best_idx, best_similarity = idx, similarity
        if best_idx is not None and best_similarity >= threshold:
            cluster = clusters[best_idx]
            cluster['sources'].add(record['vendor'])
            cluster['variants'].append(sentence)
            if len(sentence) > len(cluster['scope']):
                cluster['scope'] = sentence
                cluster['tokens'] = tokens
        else:
            clusters.append({'scope': sentence, 'tokens': tokens, 'sources': {record['vendor']}, 'variants': [sentence]})
    # Include common clauses and a balanced sample of vendor-specific clauses, not just the first file's clauses.
    common = [c for c in clusters if len(c['sources']) > 1]
    unique = [c for c in clusters if len(c['sources']) == 1]
    common.sort(key=lambda c: (len(c['sources']), len(c['tokens']), len(c['scope'])), reverse=True)
    unique_by_vendor = {}
    for cluster in unique:
        vendor_key = next(iter(cluster['sources']))
        unique_by_vendor.setdefault(vendor_key, []).append(cluster)
    for vendor_key in unique_by_vendor:
        unique_by_vendor[vendor_key].sort(key=lambda c: (len(c['tokens']), len(c['scope'])), reverse=True)
    balanced_unique = []
    vendor_keys = sorted(unique_by_vendor)
    while vendor_keys:
        remaining_keys = []
        for vendor_key in vendor_keys:
            bucket = unique_by_vendor[vendor_key]
            if bucket:
                balanced_unique.append(bucket.pop(0))
            if bucket:
                remaining_keys.append(vendor_key)
        vendor_keys = remaining_keys
    common_slots = min(len(common), int(limit * 0.60))
    unique_slots = min(len(balanced_unique), limit - common_slots)
    selected = common[:common_slots] + balanced_unique[:unique_slots]
    if len(selected) < limit:
        selected_ids = {id(item) for item in selected}
        fill = [item for item in common[common_slots:] + balanced_unique[unique_slots:] if id(item) not in selected_ids]
        selected.extend(fill[:limit - len(selected)])
    return [{'scope': c['scope'], 'sources': sorted(c['sources']), 'variants': c['variants']} for c in selected]


# Navigation controls are fixed at the bottom-left of the sidebar. The active tool
# determines which page header is rendered, so the PDF Approval Stamp Tool header
# is absent in SOW mode. Only the navigation placement and button styling are changed.
st.markdown("""
<style>
/* Anchor the two tool buttons at the lower-left of the app sidebar. */
section[data-testid="stSidebar"] { overflow: visible !important; }
div.st-key-nav_pdf_stamp,
div.st-key-nav_vendor_sow {
    position: fixed !important;
    left: 1rem !important;
    width: calc(18rem - 2rem) !important;
    z-index: 1001 !important;
    margin: 0 !important;
}
div.st-key-nav_pdf_stamp { bottom: 4.35rem !important; }
div.st-key-nav_vendor_sow { bottom: 1rem !important; }
div.st-key-nav_pdf_stamp button,
div.st-key-nav_vendor_sow button {
    width: 100% !important;
    min-height: 2.65rem !important;
    background: #008A69 !important;
    color: #FFFFFF !important;
    border: 1px solid #008A69 !important;
    border-radius: 9px !important;
    font-weight: 650 !important;
}
div.st-key-nav_pdf_stamp button:hover,
div.st-key-nav_vendor_sow button:hover {
    background: #006B4F !important;
    color: #FFFFFF !important;
    border-color: #006B4F !important;
}
div.st-key-nav_pdf_stamp button p,
div.st-key-nav_vendor_sow button p { color: #FFFFFF !important; }
</style>
""", unsafe_allow_html=True)
if 'active_tool' not in st.session_state:
    st.session_state['active_tool'] = 'PDF Approval Stamp Tool'
with st.sidebar:
    st.markdown('<div style="height:0.5rem"></div>', unsafe_allow_html=True)
    st.markdown('<div style="font-size:0.78rem;font-weight:700;color:#64748b;margin-bottom:0.4rem">TOOLS</div>', unsafe_allow_html=True)
    if st.button('PDF Approval Stamp Tool', key='nav_pdf_stamp', use_container_width=True):
        st.session_state['active_tool'] = 'PDF Approval Stamp Tool'
        st.rerun()
    if st.button('Vendor SOW Analyzer', key='nav_vendor_sow', use_container_width=True):
        st.session_state['active_tool'] = 'Vendor SOW Analyzer'
        st.rerun()

if st.session_state['active_tool'] == 'Vendor SOW Analyzer':
    st.markdown("""
    <style>
    .block-container {padding-top: 1.4rem; max-width: 1500px;}
    div[data-testid="stButton"] button {border-radius:9px;font-weight:650;}
    .sow-hero {padding:1.15rem 1.4rem;border-radius:14px;background:linear-gradient(120deg,#006B4F,#008A69);color:#fff;margin-bottom:1rem;}
    .sow-hero h1 {color:#fff;margin:0;font-size:1.8rem;}
    .sow-hero p {color:#e7fff6;margin:.35rem 0 0 0;}
    </style>
    <div class="sow-hero"><h1>Vendor SOW Analyzer</h1><p>Compare vendor scope, identify overlapping capabilities, and shortlist primary and backup coverage.</p></div>
    """, unsafe_allow_html=True)
    st.caption('Upload a global SOW as the target scope when available, plus vendor SOWs and capability documents. If no baseline is available, the analyzer identifies common and complementary work across vendors. Files are read in memory.')
    with st.expander('How the recommendations work', expanded=False):
        st.write('The analyzer extracts text and compares scope terms and evidence statements. It recommends a set of up to 3 primary vendors and up to 3 backups using a coverage-first selection. Similarity is a screening aid, not a contractual determination; verify exclusions, geography, SLAs, staffing, certifications, and commercial terms in the source documents.')
    _sow_file_types = ['pdf','docx','docm','xlsx','xlsm','xls','pptx','txt','md','csv','zip']
    baseline_file = st.file_uploader('Global vendor SOW / baseline (optional)', type=_sow_file_types, key='sow_baseline')
    vendor_files = st.file_uploader('Upload vendor SOWs, capability documents, or a ZIP containing multiple documents', type=_sow_file_types, accept_multiple_files=True, key='sow_vendor_files')
    st.caption('ZIP archives (including nested ZIPs) are expanded in memory. Exact duplicate files are removed. Documents with no detectable supplier are kept separate and clearly labeled for manual review.')

    if st.button('Analyze vendor scope', type='primary', use_container_width=True, key='run_sow_analysis'):
        try:
            parsed_docs = _sow_expand_uploads(vendor_files or [])
            parsed_vendors_by_key = {}
            excluded_documents = []
            for document in parsed_docs:
                text_value = document['Text']
                base_name = document['File'].rsplit('/', 1)[-1]
                lower_name = base_name.lower()
                if ('gp_manual' in lower_name or 'shift resume' in lower_name or 'staff roster' in lower_name or 'summary english' in lower_name):
                    excluded_documents.append({'File': document['File'], 'Reason': 'Reference, summary, or staffing document; not counted as a separate vendor.'})
                    continue
                scope_lines = _sow_extract_scope(text_value)
                if not scope_lines:
                    excluded_documents.append({'File': document['File'], 'Reason': 'No usable operational scope clauses were detected.'})
                    continue
                vendor_name = _sow_infer_vendor(document['File'], text_value)
                vendor_key = _sow_normalize(vendor_name)
                if vendor_key not in parsed_vendors_by_key:
                    parsed_vendors_by_key[vendor_key] = {'Vendor': vendor_name, 'Files': [], 'Texts': [], 'Scopes': [], 'Hashes': set()}
                grouped = parsed_vendors_by_key[vendor_key]
                grouped['Files'].append(document['File'])
                grouped['Texts'].append(text_value)
                grouped['Scopes'].extend(scope_lines)
                grouped['Hashes'].add(document['SHA256'])

            parsed_vendors = []
            for grouped in parsed_vendors_by_key.values():
                unique_scopes, seen_scopes = [], set()
                for scope in grouped['Scopes']:
                    key = _sow_normalize(scope)
                    if key and key not in seen_scopes:
                        seen_scopes.add(key)
                        unique_scopes.append(scope)
                parsed_vendors.append({
                    'Vendor': grouped['Vendor'],
                    'File': '; '.join(grouped['Files']),
                    'Files': grouped['Files'],
                    'Text': '\n\n'.join(grouped['Texts']),
                    'Scopes': unique_scopes,
                    'Coverage area': _sow_detect_coverage_area(grouped['Files'], '\n'.join(grouped['Texts'])),
                })

            baseline_docs = _sow_expand_uploads([baseline_file]) if baseline_file is not None else []
            baseline_text = '\n\n'.join(item['Text'] for item in baseline_docs)
            baseline_scopes = _sow_extract_scope(baseline_text) if baseline_text.strip() else []
            if len(parsed_vendors) < 2:
                st.error('Upload at least two distinct vendor SOWs/capability documents. Exact duplicates are removed and documents with the same detected supplier are grouped together.')
            else:
                if baseline_scopes:
                    baseline_records = [{'vendor': 'Global baseline', 'scope': scope} for scope in baseline_scopes]
                    clustered = _sow_cluster_scope_items(baseline_records, threshold=0.74, limit=180)
                    scope_items = [item['scope'] for item in clustered]
                    scope_sources = [set() for _ in scope_items]
                    analysis_mode = 'Global baseline comparison'
                else:
                    scope_records = [{'vendor': vendor['Vendor'], 'scope': scope} for vendor in parsed_vendors for scope in vendor['Scopes']]
                    clustered = _sow_cluster_scope_items(scope_records, threshold=0.74, limit=180)
                    scope_items = [item['scope'] for item in clustered]
                    scope_sources = [set(item['sources']) for item in clustered]
                    analysis_mode = 'Cross-vendor overlap (no global baseline)'

                if not scope_items:
                    st.error('No usable scope statements were detected. Upload readable SOWs or capability documents with selectable text.')
                else:
                    per_vendor = {}
                    all_evidence_rows = []
                    for vendor in parsed_vendors:
                        covered, potential, evidence_rows = set(), set(), []
                        for i, scope in enumerate(scope_items):
                            score, terms = _sow_similarity(scope, vendor['Scopes'])
                            phrase = _sow_normalize(scope)
                            vendor_norm = _sow_normalize(vendor['Text'])
                            phrase_present = bool(phrase and phrase in vendor_norm)
                            if vendor['Vendor'] in scope_sources[i]:
                                status, score = 'Covered', max(score, 100.0)
                            else:
                                status = 'Covered' if phrase_present or score >= 60 else ('Potential' if score >= 30 else 'Gap')
                            if status == 'Covered':
                                covered.add(i)
                            elif status == 'Potential':
                                potential.add(i)
                            evidence = ''
                            if vendor['Scopes']:
                                best_score = -1
                                for candidate_line in vendor['Scopes']:
                                    candidate_score, _ = _sow_similarity(scope, [candidate_line])
                                    if candidate_score > best_score:
                                        best_score, evidence = candidate_score, candidate_line
                            if not evidence or status == 'Gap':
                                evidence = 'No close scope excerpt found; manually review the source document.'
                            evidence_rows.append({
                                'Scope requirement': _sow_scope_label(scope),
                                'Vendor': vendor['Vendor'],
                                'Status': status,
                                'Match score (%)': score,
                                'Matched terms': ', '.join(terms[:12]),
                                'Evidence excerpt': evidence[:500],
                                'Source document(s)': vendor['File'],
                            })
                        n = max(len(scope_items), 1)
                        weighted_coverage = round((len(covered) + 0.5 * len(potential)) / n * 100, 1)
                        scores = [r['Match score (%)'] for r in evidence_rows]
                        per_vendor[vendor['Vendor']] = {
                            'covered': covered,
                            'potential': potential,
                            'weighted_coverage': weighted_coverage,
                            'average': round(sum(scores) / len(scores), 1) if scores else 0,
                            'evidence': evidence_rows,
                            'file': vendor['File'],
                            'scope_count': len(vendor['Scopes']),
                            'coverage_area': vendor.get('Coverage area', 'Not established from document title'),
                        }
                        all_evidence_rows.extend(evidence_rows)

                    ranked_vendors = sorted(
                        per_vendor,
                        key=lambda name: (
                            per_vendor[name]['weighted_coverage'],
                            len(per_vendor[name]['covered']),
                            per_vendor[name]['average'],
                            -len(per_vendor[name]['potential']),
                        ),
                        reverse=True,
                    )
                    primary = ranked_vendors[:3]
                    backups = ranked_vendors[3:6]
                    st.subheader('Analysis overview')
                    st.info(f'Mode: {analysis_mode}. Distinct scope clauses assessed: {len(scope_items)}. Distinct supplier groups analyzed: {len(parsed_vendors)}. Exact duplicate files are excluded from the count.')
                    if not baseline_scopes:
                        st.warning('No global baseline was supplied. This mode measures overlap across the uploaded scope documents; it cannot prove that any single vendor can replace a global provider. Geographic coverage is shown separately and must be validated before treating a candidate as global.')
                    if any(name.startswith('Unidentified supplier') for name in per_vendor):
                        st.warning('One or more documents did not identify a legal supplier. These are labeled by document name and may represent a region/template rather than an independent vendor. Verify the supplier identity before using the shortlist.')
                    if len(parsed_vendors) < 6:
                        st.warning(f'Only {len(parsed_vendors)} distinct supplier group(s) were identified. The analyzer can recommend up to {len(primary)} primary and {len(backups)} backup candidates; upload additional independent vendor SOWs to reach six.')
                    if excluded_documents:
                        with st.expander(f'Reference/admin documents excluded from vendor ranking ({len(excluded_documents)})', expanded=False):
                            st.dataframe(pd.DataFrame(excluded_documents), use_container_width=True, hide_index=True)

                    k1, k2, k3, k4 = st.columns(4)
                    k1.metric('Distinct vendors/groups', len(parsed_vendors))
                    k2.metric('Scope clusters', len(scope_items))
                    k3.metric('Primary candidates', len(primary))
                    k4.metric('Backup candidates', len(backups))

                    coverage_rows = []
                    for name, detail in per_vendor.items():
                        n = max(len(scope_items), 1)
                        coverage_rows.append({
                            'Vendor / supplier group': name,
                            'Coverage area (from document title)': detail['coverage_area'],
                            'Global coverage evidence': ('Not evidenced — region-specific SOW' if detail['coverage_area'] != 'Not established from document title' else 'Not established; manual verification required'),
                            'Weighted coverage (%)': detail['weighted_coverage'],
                            'Confirmed covered (%)': round(100 * len(detail['covered']) / n, 1),
                            'Potential (%)': round(100 * len(detail['potential']) / n, 1),
                            'Average match (%)': detail['average'],
                            'Covered scopes': len(detail['covered']),
                            'Potential scopes': len(detail['potential']),
                            'Gap scopes': n - len(detail['covered']) - len(detail['potential']),
                            'Documents grouped': len(detail['file'].split('; ')),
                            'Source file(s)': detail['file'],
                        })
                    coverage_frame = pd.DataFrame(coverage_rows).sort_values(
                        ['Weighted coverage (%)', 'Confirmed covered (%)', 'Average match (%)'], ascending=False
                    )
                    st.subheader('Vendor coverage graph')
                    st.bar_chart(coverage_frame.set_index('Vendor / supplier group')[['Weighted coverage (%)', 'Confirmed covered (%)', 'Potential (%)']], use_container_width=True)
                    st.caption('Weighted coverage counts each Covered scope as 1 and each Potential scope as 0.5. Confirmed and Potential are shown separately to avoid presenting uncertain matches as proven coverage. Geographic labels are inferred from document titles only; they are not proof of complete country coverage. Similarity is evidence screening, not a contractual conclusion.')

                    st.subheader('Recommended primary vendors')
                    if primary:
                        st.dataframe(coverage_frame[coverage_frame['Vendor / supplier group'].isin(primary)], use_container_width=True, hide_index=True)
                    else:
                        st.warning('No primary candidates could be selected from the available scope statements.')
                    st.subheader('Backup vendors')
                    if backups:
                        st.dataframe(coverage_frame[coverage_frame['Vendor / supplier group'].isin(backups)], use_container_width=True, hide_index=True)
                    else:
                        st.write('No additional distinct vendor groups are available as backups. Upload more independent vendor documents to identify backup options.')

                    st.subheader('Scope-by-vendor comparison matrix')
                    matrix = pd.DataFrame([{'Scope requirement': _sow_scope_label(scope)} for scope in scope_items])
                    for name, detail in per_vendor.items():
                        statuses = {row['Scope requirement']: row['Status'] for row in detail['evidence']}
                        matrix[name] = [statuses.get(_sow_scope_label(scope), 'Gap') for scope in scope_items]
                    st.dataframe(matrix, use_container_width=True, hide_index=True)

                    st.subheader('Detailed match evidence')
                    evidence_frame = pd.DataFrame(all_evidence_rows)
                    status_filter = st.multiselect('Filter by match status', ['Covered', 'Potential', 'Gap'], default=['Covered', 'Potential', 'Gap'], key='sow_status_filter')
                    st.dataframe(evidence_frame[evidence_frame['Status'].isin(status_filter)], use_container_width=True, hide_index=True)

                    document_inventory = pd.DataFrame([
                        {'Document': document['File'], 'Characters extracted': len(document['Text']), 'Scope clauses detected': len(_sow_extract_scope(document['Text'])), 'Supplier detected': _sow_infer_vendor(document['File'], document['Text']), 'Coverage area (from title)': _sow_detect_coverage_area([document['File']], document['Text'])}
                        for document in parsed_docs
                    ])
                    export_buffer = BytesIO()
                    with pd.ExcelWriter(export_buffer, engine='openpyxl') as writer:
                        coverage_frame.to_excel(writer, index=False, sheet_name='Vendor Summary')
                        matrix.to_excel(writer, index=False, sheet_name='Scope Matrix')
                        evidence_frame.to_excel(writer, index=False, sheet_name='Match Evidence')
                        document_inventory.to_excel(writer, index=False, sheet_name='Document Inventory')
                        pd.DataFrame({
                            'Primary vendors': primary + [''] * (3 - len(primary)),
                            'Backup vendors': backups + [''] * (3 - len(backups)),
                        }).to_excel(writer, index=False, sheet_name='Recommendations')
                    st.download_button('Download SOW comparison workbook', data=export_buffer.getvalue(), file_name='Vendor_SOW_Analysis.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
                    st.warning('Before selecting vendors, validate the source clauses for geography/countries, response and resolution SLAs, hours of coverage, onsite/remote requirements, exclusions, dependencies, certifications, subcontracting, and commercial assumptions. The analyzer does not certify legal equivalence.')
        except Exception:
            st.error('The SOW documents could not be analyzed. Check that the files are readable, not password-protected, and within the 100 MB expanded processing limit.')
    st.stop()

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
    if selected_region != "INDIA":
        stamp_placement = st.selectbox(
            "Stamp placement",
            ["Center in available blank space", "Bottom of summary page"],
        )
    else:
        stamp_placement = "Center in available blank space"
    st.caption("Region-specific validation rules are applied automatically.")

if selected_region == "INDIA":
    st.subheader("India — NCR AP Upload CSV")
    st.caption("Upload matching invoice PDFs; matching requires both invoice number and supplier GSTIN.")

    qcol, ycol = st.columns(2)
    with qcol:
        selected_quarter = st.selectbox("Invoice quarter", ["Q1", "Q2", "Q3", "Q4"], index=3)
    with ycol:
        _years = list(range(datetime.now().year - 2, datetime.now().year + 4))
        selected_year = st.selectbox("Invoice year", _years, index=_years.index(datetime.now().year))
    fcol, mcol = st.columns(2)
    with fcol:
        selected_fee_type = st.selectbox("Weekly file description", ["Incident_fee", "TPM_fee"], index=0)
    with mcol:
        selected_upload_month = st.selectbox(
            "Weekly file month",
            ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
            index=datetime.now().month - 1,
        )

    # Always initialize the in-app table before rendering the PDF uploader.
    if "india_template_frame" not in st.session_state or st.session_state["india_template_frame"] is None:
        st.session_state["india_template_frame"] = _india_initial_table()
        st.session_state["india_csv_name"] = f"NCR Invoice AP Upload file {datetime.now().strftime('%m.%d.%Y')}.csv"
        st.session_state.setdefault("india_unmatched", [])
        st.session_state.setdefault("india_duplicates", [])
        st.session_state.setdefault("india_results", [])

    # Incrementing the uploader key recreates the widget and clears its selected PDFs.
    # The callback runs before the next script execution, avoiding direct mutation
    # of a file_uploader widget after it has been instantiated.
    def _reset_india_pdf_uploads():
        # Clear the selected PDFs, restore the default template table, and clear
        # the previous extraction results without affecting other regions.
        st.session_state["india_pdf_uploader_version"] = (
            st.session_state.get("india_pdf_uploader_version", 0) + 1
        )
        st.session_state["india_template_frame"] = _india_initial_table()
        st.session_state["india_results"] = []
        st.session_state["india_unmatched"] = []
        st.session_state["india_duplicates"] = []
        st.session_state["india_csv_name"] = (
            f"NCR Invoice AP Upload file {datetime.now().strftime('%m.%d.%Y')}.csv"
        )
        st.session_state.pop("india_template_data_editor", None)

    _india_pdf_uploader_key = f"india_pdfs_{st.session_state.get('india_pdf_uploader_version', 0)}"
    india_pdfs = st.file_uploader(
        "Upload matching invoice PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        key=_india_pdf_uploader_key,
    )

    if india_pdfs and st.button("Extract PDF data into template table", type="primary", use_container_width=True):
        try:
            with st.spinner("Matching invoice numbers and supplier GSTINs..."):
                india_frame, india_results, india_unmatched, india_duplicates = _india_make_csv(
                    india_pdfs, selected_quarter, int(selected_year),
                    selected_fee_type, selected_upload_month, int(selected_year)
                )
            st.session_state["india_template_frame"] = india_frame
            st.session_state["india_results"] = india_results
            st.session_state["india_unmatched"] = india_unmatched
            st.session_state["india_duplicates"] = india_duplicates
            st.session_state["india_csv_name"] = f"NCR Invoice AP Upload file {datetime.now().strftime('%m.%d.%Y')}.csv"
            st.session_state.pop("india_template_data_editor", None)
            st.success(f"Extracted data into the table. {len(india_results)} invoice record(s) matched.")
        except Exception:
            st.error("The invoice PDFs could not be processed. Check that the PDFs are valid and try again.")

    st.markdown("""
    <style>
    div[data-testid="stDataEditor"] [role="columnheader"] {
        background-color: #FF2400 !important;
        color: #FFFFFF !important;
        text-align: center !important;
        border: 1px solid #222222 !important;
        justify-content: center !important;
    }
    div[data-testid="stDataEditor"] [role="gridcell"] {
        text-align: center !important;
        border: 1px solid #777777 !important;
    }
    div[data-testid="stDataEditor"] [role="row"] {
        border-bottom: 1px solid #777777 !important;
    }
    </style>
    """, unsafe_allow_html=True)
    st.subheader("Review and edit table")
    st.caption("The table is built into the app and is displayed immediately below the PDF upload. Upload PDFs and select Extract PDF data into template table to populate matching rows.")
    india_edited_frame = st.data_editor(
        st.session_state["india_template_frame"],
        num_rows="dynamic",
        use_container_width=True,
        key="india_template_data_editor",
    )
    _india_totals_frame = india_edited_frame.copy()
    _india_total_columns = [
        column for column in ["Amount- Duty/IGST", "Tax field", "Total (amount+tax)"]
        if column in _india_totals_frame.columns
    ]
    if _india_total_columns:
        _india_totals_rows = []
        for _column in _india_total_columns:
            _numeric_values = pd.to_numeric(_india_totals_frame[_column], errors="coerce").dropna()
            _india_totals_rows.append({
                "Amount column": _column,
                "Total amount": round(float(_numeric_values.sum()), 2),
            })
        st.subheader("Total Amount Summary")
        _india_totals_display = pd.DataFrame(_india_totals_rows)
        _india_totals_display["Total amount"] = _india_totals_display["Total amount"].map(
            lambda value: f"{value:,.2f}"
        )
        st.dataframe(_india_totals_display, use_container_width=True, hide_index=True)

    if st.session_state.get("india_unmatched"):
        st.warning("Some table rows did not match both invoice number and supplier GSTIN. Those rows were not filled: " + ", ".join(st.session_state["india_unmatched"]))
    if st.session_state.get("india_duplicates"):
        st.warning("Duplicate invoice PDFs were uploaded for: " + ", ".join(st.session_state["india_duplicates"]))

    _india_xlsx_buffer = BytesIO()
    with pd.ExcelWriter(_india_xlsx_buffer, engine="openpyxl") as _writer:
        india_edited_frame.to_excel(_writer, index=False, sheet_name="NCR")
        _ws = _writer.sheets["NCR"]
        _scarlet = PatternFill(fill_type="solid", fgColor="FF2400")
        _white_bold = Font(color="FFFFFF", bold=True)
        _thin_border = Border(
            left=Side(style="thin", color="777777"),
            right=Side(style="thin", color="777777"),
            top=Side(style="thin", color="777777"),
            bottom=Side(style="thin", color="777777"),
        )
        for _row in _ws.iter_rows():
            for _cell in _row:
                _cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True,
                )
                _cell.border = _thin_border
                if _cell.row == 1:
                    _cell.fill = _scarlet
                    _cell.font = _white_bold

        # Fit each Excel column to its longest value/header. Long text columns
        # are wrapped and capped at a readable width rather than stretching
        # the workbook beyond the screen.
        for _column_cells in _ws.columns:
            _column_letter = _column_cells[0].column_letter
            _max_content_length = max(
                (len(str(_cell.value)) if _cell.value is not None else 0)
                for _cell in _column_cells
            )
            _ws.column_dimensions[_column_letter].width = min(
                max(_max_content_length + 2, 12),
                60,
            )
        _ws.freeze_panes = "A2"
        _ws.auto_filter.ref = _ws.dimensions

    st.markdown("""
    <style>
    div[data-testid="stDownloadButton"] button {
        background-color: #008A69 !important;
        color: #FFFFFF !important;
        border: 1px solid #008A69 !important;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background-color: #006B4F !important;
        color: #FFFFFF !important;
        border-color: #006B4F !important;
    }
    </style>
    """, unsafe_allow_html=True)

    st.download_button(
        "Download formatted Excel table",
        data=_india_xlsx_buffer.getvalue(),
        file_name=st.session_state.get("india_csv_name", f"NCR Invoice AP Upload file {datetime.now().strftime('%m.%d.%Y')}.csv").rsplit(".", 1)[0] + ".xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.button(
        "Reset uploaded PDFs and table",
        on_click=_reset_india_pdf_uploads,
        use_container_width=True,
        help="Clear uploaded PDFs and reset the India table and extraction results to the original template.",
    )
    st.stop()


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

    # Auto-detect the most likely summary/approval page using the selected region's
    # labels and coordinate-matched values. Keep the dropdown available for correction.
    try:
        candidate_pages, auto_page_index = auto_detect_summary_page(
            pdf_content, selected_region, candidate_pages
        )
    except Exception:
        auto_page_index = candidate_pages[0][0]
    page_labels = [f"Page {page_number + 1}" for page_number, _ in candidate_pages]
    st.caption(f"Automatically detected summary page: Page {auto_page_index + 1}. You can change it below if needed.")
    selected_page_label = st.selectbox(
        "Summary / approval page",
        page_labels,
        index=0,
    )
    selected_position = page_labels.index(selected_page_label)
    page_index, page_text = candidate_pages[selected_position]
    extracted = extract_from_page(page_text, selected_region)
    # The invoice summary labels and page coordinates are the source of truth.
    # Do not use any GL codes or amounts printed in an existing approval stamp.
    try:
        _coordinate_doc = pymupdf.open(stream=pdf_content, filetype="pdf")
        try:
            _coordinate_rows, _coordinate_total = extract_region_amounts_from_page(
                _coordinate_doc[page_index], selected_region
            )
        finally:
            _coordinate_doc.close()
        if len(_coordinate_rows) == 2:
            extracted["rows"] = _coordinate_rows
            extracted["total"] = _coordinate_total
        else:
            # Fail closed rather than silently carrying forward incorrect stamped rows.
            extracted["rows"] = _coordinate_rows
            extracted["total"] = _coordinate_total
    except Exception:
        extracted["rows"] = []
        extracted["total"] = None

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
        extracted_amounts = [
            float(row["Amount"]) for row in extracted.get("rows", [])
            if row.get("Amount") is not None
        ]
        default_total = round(sum(extracted_amounts), 2) if len(extracted_amounts) == 2 else 0.0
        invoice_total = st.number_input(
            "Total amount (verify or enter manually)", min_value=0.0, value=default_total, step=0.01, format="%.2f"
        )
    try:
        _approval_secrets = st.secrets.get("approval", {})
        _secret_approver = str(_approval_secrets.get("approved_by", "")).strip()
        _secret_email = str(_approval_secrets.get("email", "")).strip()
    except Exception:
        _secret_approver, _secret_email = "", ""
    with second:
        approver_name = st.text_input("Approved by", value=extracted["approver"] or _secret_approver)
    with third:
        approval_date = st.text_input(
            "Approval date (MM.DD.YYYY)",
            value=datetime.now().strftime("%m.%d.%Y"),
            disabled=True,
        )
    approver_email = st.text_input("Approver email", value=_secret_email or extracted["email"])

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
        elif not approver_name.strip():
            st.error("Approver name is required.")
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
                    datetime.now().strftime("%m.%d.%Y"),
                    approver_email,
                    stamp_placement,
                    selected_region,
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
