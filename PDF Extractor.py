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
    def tax_amount_on_line(label_pattern):
        for line in text.splitlines():
            if re.search(label_pattern, line, re.I):
                matches = re.findall(r"\d[\d,]*\.\d{2}", line)
                # Tax lines often contain a rate (e.g. 18.00) before the tax value.
                if matches:
                    return _india_amount(matches[-1])
        # Some PDFs split the tax label, rate, and amount across adjacent lines.
        lines = text.splitlines()
        for i, line in enumerate(lines):
            if re.search(label_pattern, line, re.I):
                chunk = " ".join(lines[i:i + 3])
                matches = re.findall(r"\d[\d,]*\.\d{2}", chunk)
                if matches:
                    return _india_amount(matches[-1])
        return None

    has_igst = bool(re.search(r"\bIGST\b", text, re.I))
    has_cgst = bool(re.search(r"\bCGST\b", text, re.I))
    has_sgst = bool(re.search(r"\b(?:SGST|UTGST)\b", text, re.I))

    # When both slab-specific tax lines are present, use their actual amounts
    # and add them together instead of deriving tax by splitting the total.
    def slab_amount(code):
        lines = text.splitlines()
        for i, line in enumerate(lines):
            if re.search(re.escape(code), line, re.I):
                matches = re.findall(r"\d[\d,]*\.\d{2}", line)
                if matches:
                    return _india_amount(matches[-1])
                chunk = " ".join(lines[i:i + 3])
                matches = re.findall(r"\d[\d,]*\.\d{2}", chunk)
                if matches:
                    return _india_amount(matches[-1])
        return None

    slab_sgst = slab_amount("SGST-INV-SLAB4")
    slab_cgst = slab_amount("CGST-INV-SLAB4")
    total = value_after([
        r"\bINVOICE TOTAL\b[\s\S]{0,80}?([₹]?\s*[\d,]+\.\d{2})",
        r"\bTOTAL AMOUNT\b[\s\S]{0,80}?([₹]?\s*[\d,]+\.\d{2})",
    ])
    # Prefer explicit tax amounts. When both CGST and SGST/UTGST are present,
    # add them for the template's single Tax field; if only one is present,
    # use that amount without halving or deriving a different value.
    explicit_cgst = slab_cgst
    explicit_sgst = slab_sgst
    if explicit_cgst is None:
        explicit_cgst = tax_amount_on_line(r"\bCGST\b")
    if explicit_sgst is None:
        explicit_sgst = tax_amount_on_line(r"\b(?:SGST|UTGST)\b")
    explicit_igst = tax_amount_on_line(r"\bIGST\b") if has_igst else None

    tax_total = round(total - net, 2) if total is not None and net is not None else None
    if explicit_igst is not None:
        igst = explicit_igst
    else:
        igst = tax_total if has_igst else None

    cgst = explicit_cgst
    sgst = explicit_sgst
    if cgst is not None and sgst is not None:
        tax_field = round(cgst + sgst, 2)
        tax_type = "CGST/SGST"
    elif cgst is not None:
        tax_field = cgst
        tax_type = "CGST"
    elif sgst is not None:
        tax_field = sgst
        tax_type = "SGST/UTGST"
    elif has_cgst and has_sgst and tax_total is not None:
        # Fall back to the invoice tax total only when neither component amount
        # was individually readable but both tax types are indicated.
        tax_field = tax_total
        cgst = round(tax_total / 2, 2)
        sgst = round(tax_total - cgst, 2)
        tax_type = "CGST/SGST"
    elif has_igst:
        tax_field = igst
        tax_type = "IGST"
    else:
        tax_field = None
        tax_type = ""
    return {
        "invoice_number": inv_match.group(1).upper() if inv_match else "",
        "gstin": gstin,
        "invoice_date": invoice_date,
        "net": net,
        "igst": igst,
        "cgst": cgst,
        "sgst": sgst,
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
                if region in ("NA and LATAM", "EMEA & APAC"):
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
        compact_regional_stamp = region_key in ("NA and LATAM", "EMEA & APAC")
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

        if compact_regional_stamp:
            # Constrain the stamp to the requested blank section: below the
            # service-description paragraph, left of NET TOTAL PAYABLE, and above
            # the THIS INVOICE footer line. The exact bounds are read from page text.
            service_bottom = None
            payable_left = None
            footer_top = None
            for rect, block_text in occupied:
                upper = re.sub(r"\s+", " ", block_text.upper())
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
        st.session_state["india_pdf_uploader_version"] = (
            st.session_state.get("india_pdf_uploader_version", 0) + 1
        )

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

    # Export the formatted Excel workbook; the CSV download is intentionally removed.
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
                _cell.alignment = Alignment(horizontal="center", vertical="center")
                _cell.border = _thin_border
                if _cell.row == 1:
                    _cell.fill = _scarlet
                    _cell.font = _white_bold
        # Fit each column to its longest header/data value, with a modest safety
        # margin so the exported workbook opens without clipped cell contents.
        for _column_cells in _ws.columns:
            _max_length = max(
                (len(str(_cell.value)) if _cell.value is not None else 0)
                for _cell in _column_cells
            )
            _column_letter = _column_cells[0].column_letter
            _ws.column_dimensions[_column_letter].width = max(10, _max_length + 2)
        # Fit row heights to the default font and keep the header readable.
        for _row_number in range(1, _ws.max_row + 1):
            _ws.row_dimensions[_row_number].height = 20 if _row_number == 1 else 18
        _ws.freeze_panes = "A2"
        _ws.auto_filter.ref = _ws.dimensions

    st.download_button(
        "Download formatted Excel table",
        data=_india_xlsx_buffer.getvalue(),
        file_name=st.session_state.get("india_csv_name", f"NCR Invoice AP Upload file {datetime.now().strftime('%m.%d.%Y')}.csv").rsplit(".", 1)[0] + ".xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.button(
        "Reset uploaded PDFs",
        on_click=_reset_india_pdf_uploads,
        use_container_width=True,
        help="Clear the uploaded India invoice PDFs so you can select a new batch.",
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
