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
from datetime import datetime

import streamlit as st
import pymupdf
import pandas as pd
import pytesseract
from PIL import Image
from io import BytesIO
from pathlib import Path

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
    total = value_after([
        r"\bINVOICE TOTAL\b[\s\S]{0,80}?([₹]?\s*[\d,]+\.\d{2})",
        r"\bTOTAL AMOUNT\b[\s\S]{0,80}?([₹]?\s*[\d,]+\.\d{2})",
    ])
    # Summary PDFs may put the tax label, rate, and tax amount in separate
    # text blocks. Derive total tax from invoice total minus subtotal, then
    # use the template's single-tax-field convention for intra-state invoices.
    tax_total = round(total - net, 2) if total is not None and net is not None else None
    igst = tax_total if has_igst else None
    cgst = round(tax_total / 2, 2) if has_cgst and has_sgst and tax_total is not None else None
    sgst = cgst if has_cgst and has_sgst else None
    tax_field = igst if has_igst else cgst
    tax_type = "IGST" if has_igst else ("CGST/SGST" if has_cgst and has_sgst else "")
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


def _india_make_csv(template_bytes: bytes, pdf_uploads, quarter: str, year: int, fee_type: str, upload_month: str, upload_year: int):
    """Fill blank AP upload template fields using invoice number + supplier GSTIN."""
    try:
        frame = pd.read_excel(BytesIO(template_bytes), sheet_name="NCR", dtype=object)
    except Exception:
        frame = pd.read_excel(BytesIO(template_bytes), dtype=object)
    frame.columns = [_india_clean_text(c) for c in frame.columns]
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
    st.caption("The NCR AP Excel template is built into this script. Upload the matching invoice PDFs; matching requires both invoice number and supplier GSTIN.")
    qcol, ycol = st.columns(2)
    with qcol:
        selected_quarter = st.selectbox("Invoice quarter", ["Q1", "Q2", "Q3", "Q4"], index=3)
    with ycol:
        selected_year = st.selectbox("Invoice year", list(range(datetime.now().year - 2, datetime.now().year + 4)), index=2)
    fcol, mcol = st.columns(2)
    with fcol:
        selected_fee_type = st.selectbox("Weekly file description", ["Incident_fee", "TPM_fee"], index=0)
    with mcol:
        selected_upload_month = st.selectbox("Weekly file month", ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"], index=datetime.now().month - 1)
    # The Excel template is embedded in this script so no separate template upload/file is required.
    _DEFAULT_TEMPLATE_B64 = "UEsDBBQABgAIAAAAIQDVdhn1fQEAAAkGAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADEVMluwjAQvVfqP0S+VsTAoaoqAocuxxYJ+gEmHhKLxLY8A4W/78QsqioWRSD1kiix5232zGC0rqtkBQGNs5nopV2RgM2dNrbIxNf0vfMkEiRltaqchUxsAMVoeH83mG48YMLVFjNREvlnKTEvoVaYOg+WV+Yu1Ir4MxTSq3yhCpD9bvdR5s4SWOpQgyGGg1eYq2VFyduaf2+VzIwVyct2X0OVCeV9ZXJFLFSurP5D0nHzuclBu3xZM3SKPoDSWAJQXaU+GGYMEyBiYyjkUc4AFbYj3blKuTIKw9J4fGDrJxialdOudnWffBzBaEjGKtCHqtm7XFfy24XFzLlFeh6kbTQxorRWxu51n+GPm1HGV+/GQhp/Ebiljv4/6SC+6yDj8/ooIswF40ibCvDWxx9BLzGXKoCeEHdRcXMBv7HP6eDWHgfnkadHgPYp7Fu1qe54BoJABg7NeuzSHxh59FwdOzSzTYNuy50vkVx9Nf0W5gi5jIN8+AMAAP//AwBQSwMEFAAGAAgAAAAhABNevmUCAQAA3wIAAAsACAJfcmVscy8ucmVscyCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACskk1LAzEQhu+C/yHMvTvbKiLSbC9F6E1k/QExmf1gN5mQpLr990ZBdKG2Hnqcr3eeeZn1ZrKjeKMQe3YSlkUJgpxm07tWwkv9uLgHEZNyRo3sSMKBImyq66v1M40q5aHY9T6KrOKihC4l/4AYdUdWxYI9uVxpOFiVchha9EoPqiVcleUdht8aUM00xc5ICDtzA6I++Lz5vDY3Ta9py3pvyaUjK5CmRM6QWfiQ2ULq8zWiVqGlJMGwfsrpiMr7ImMDHida/Z/o72vRUlJGJYWaA53m+ew4BbS8pEVzE3/cmUZ85zC8Mg+nWG4vyaL3MbE9Y85XzzcSzt6y+gAAAP//AwBQSwMEFAAGAAgAAAAhADhRZ/k9AwAAiAcAAA8AAAB4bC93b3JrYm9vay54bWykVdFumzAUfZ+0f2BWXykYCElQydRBolXqqqrrtpdKlQNOsWpsZpuGatq/7xqSNFleug4lNuZa555jn2uffexq7jxRpZkUKcKnPnKoKGTJxEOKvt0u3AlytCGiJFwKmqJnqtHH2ft3Z2upHpdSPjoAIHSKKmOaxPN0UdGa6FPZUAGRlVQ1MTBUD55uFCWlrig1NfcC34+9mjCBBoREvQZDrlasoLks2poKM4AoyokB+rpijd6i1cVr4GqiHtvGLWTdAMSScWaee1Dk1EVy8SCkIksOsjs8cjoFvxj+2Icm2GaC0FGqmhVKarkypwDtDaSP9GPfw/hgCbrjNXgdUuQp+sTsHu5YqfiNrOIdVvwChv3/RsNgrd4rCSzeG9FGO24Bmp2tGKffB+s6pGmuSG13iiOHE23mJTO0TNEYhnJNXz6AKtU2n1rGIRpEYRghb7az87VySroiLTe3YOQtPFRGHE+DkZ0JxjjnhipBDM2kMODDja7/9VyPnVUSHO7c0J8tUxQKC/wFWqElRUKW+pqYymkVT1GW3H3TIP+OUF2Qu1yuBZdQYHd7ziTHZfAP3iSFFeyB4oHV8P63eiCnkq3/ro1y4P0iv4Q9+EqeYEdg38tNwV7AkuPwXhQqwfe/Yn+cT6bh3I2iPHKjxQS7E4wDN5rHeZZN5lmGp79BjIqTQpLWVJvNttApimBnj0JfSLeNYD9pWflC45e/eVzb/9VsY7+tYHusfWd0rV9sYYdO94OJUq5T5OIARD0fDtd98AcrTQW+8schTBm+fabsoQLGGOPYFoEKLLMUHTDKB0YLeFzbHDDy9ij1ByhQ63tH9Ka/ym7gmLYnq13hEBye2ATqosRW0P7Ur3YW3psNB9ludtDv9zYDFAITtLR1Bfn2Rpus9x0X9en9gtlyyIkhS6KpLbeC8D6NJQOCK1aW1N4uaAZEP5ycn+Dk5PYEn3l7mOCtw3wAUkA12q63zRT7wdTyo5251KbvoRAYrGOOo6kfzs/dMMzASOPF2J0s/JEbRuMoG0Wf5tgfWyPZmyrp+Lp4etv5E0Te9trL9q+MjS9tmVrwZHOfOpqaTciKt6UE3Ie2V7BDm/0BAAD//wMAUEsDBBQABgAIAAAAIQBKqaZh+gAAAEcDAAAaAAgBeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHMgogQBKKAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAC8ks1qxDAMhO+FvoPRvXGS/lDKOnsphb222wcwsRKHTWxjqT95+5qU7jawpJfQoyQ08zHMZvs59OIdI3XeKSiyHAS62pvOtQpe909X9yCItTO69w4VjEiwrS4vNs/Ya05PZLtAIqk4UmCZw4OUVFscNGU+oEuXxsdBcxpjK4OuD7pFWeb5nYy/NaCaaYqdURB35hrEfgzJ+W9t3zRdjY++fhvQ8RkLyYkLk6COLbKCafxeFlkCBXmeoVyT4cPHA1lEPnEcVySnS7kEU/wzzGIyt2vCkNURzQvHVD46pTNbLyVzsyoMj33q+rErNM0/9nJW/+oLAAD//wMAUEsDBBQABgAIAAAAIQDhe7tTjAcAADAjAAAYAAAAeGwvd29ya3NoZWV0cy9zaGVldDEueG1spJRdb9owFIbvJ+0/RL4niUM+EbRqKWjdxTSt3XZtHANW45jZptBO++87DjEgMSHCJMI5SeznfL3O8HYrKu+VKc1lPULYD5HHaipLXi9G6PvztJcjTxtSl6SSNRuhN6bR7c3HD8ONVC96yZjxgFDrEVoasxoEgaZLJoj25YrV8GYulSAGbtUi0CvFSNlsElUQhWEaCMJrtCMM1CUMOZ9zyh4kXQtWmx1EsYoYyF8v+Uo7mqCX4ARRL+tVj0qxAsSMV9y8NVDkCTp4XNRSkVkFdW9xTKi3VfCL4Oq7MM3zk0iCUyW1nBsfyMEu59Pyi6AICN2TTuu/CIPjQLFXbgd4QEXXpYSTPSs6wPpXwtI9zLZLDda8HKHfMS7ww2Qy7uXxXdGLk/txr8jGk16WFOn9JA+nRXT3B90MSw4TtlV5is1H6A4PnnGIgpthI6AfnG30ke8ZMntiFaOGQRCMvHcpxRMlFftiFVjBsxC0bVU7k/LFbn+EhSEE0s02G4hQw1/ZmFWw/DMw9K8mNLgQNtjHPfZdDtNG51+VV7I5WVfmm9x8YnyxNBA4geqtfAbl2wPTFHQLgf0osVQqK0DAvye4PYCgO7Jt7IaXZgle7sdRkuUY1nszps2UWyby6FobKX62q1rWjgKzayhgHSX1cRymHRgwsoYBtmXkfpKF/S55xC0D7H9UA3U3mYBtKbGfJ0mc5tnlLUlbCNgWEoXdKVlLAfsvypmRwGe0KQKsa0XUuZ9FCwHrIGF3dcApakVmD8ROP0m/qz7wXqrguGzwnnKmFdjJ0zrtzuLQi3M7nSjxQZU47T5G7IRpnTaDfnRFJ5007RF3PSiuyMepEx/kmRylc64lTpL4oEkcdR6mkyc+0ucVpww7hVrHNSS7VOdB8zH8CwAA//8AAAD//5SZ7U7bQBBFXyXyAxB/5gOFSC0BG1wMGL9AlEbiD1CRlLZvX5vNmtl7xxvlF2XmzAzenN3Y7mL3vN3uV+v9erl4f/szer8IomC0+7V+3bX/Oo/D9peLIA1Gm9+7/dvL9dv7y3r/yTy3P+L4LAtGf6N0vTn/+W+13W22r204PIuD5WLTNfvWdbsIJsGoje/a6McyXIw/lovx5kB8ZyLKXORSQVxixUTsEldMJC5xrfSIXCRXELicgpHUbXLDBFzvLRMTt0fJxNQlfhgiEQs/c4k77jF3iUpZd7jcewWBRXswyEz8JRF8OI8GmUoErrhWELjkJwWBa24U5Ouix+0O6LdBDNtAszzrNe/odj98Ch4nM3RcpidhGM6nCdh36SBxGkbQY2WALBgfNs6VrfjaWhEqjTU5Bgo5NoWRNyY570fe8sgMPspSNszCSYJe2ry9jDsMVDwkheu6ZySGv+NBWR3YiI8GaZXoz6YYBtUKAl2eFAT2c2OR7vibz2fTaMC5drc6R6/fuY72OCfTA845iOacAYRztsLjHNbkGCjkWHLOJIVzPDKDE6aUDTXnbL53DgMVDyHnGCHnlNVB5wziOIcnnYKgcwoC+6yxyDHn2u/2E5zraI9zMj3gnINozhlAOGcrPM5hTY6BQo4l50xSOMcjMzgbS9lQc87me+cwUPEQco4Rck5ZHXTOII5z8L1YKwg6pyBwv9BY5Jhz7U3jCc51tMc5mR5wzkE05wwgnLMVHuewJsdAIceScyYpnOORKaxvKRtqztl87xwGKmUIfrcyQs4pq4POGUQ6l8AOqhmJ0TmlC5z9jUWOOdc+gpzgXEd7nJPpAeccRHPOAMI5W+FxDmtyDBRyLDlnksI5HpnCyVDKhppzNt87h4FKGYLOMULOMZLgU4RBHOfwfo4Rck7pAlo2FjnmXPtAc4JzHe1xTqYHnHMQzTkDCOdshcc5rMkxUMix5JxJCud4ZAr3LqVsqDln871zGKiUIegcI+Scsjp4zhnEcQ7u/mtGyDmlC9wVNhY55lz7nH2Ccx3tcU6mB5xzEM05AwjnbIXHOazJMVDIseScSQrneGSKb1RkQ805m++dw0ClDEHnGCHnlNVB5wziOAc7qGaEnFO64NsSixxzbn6Scx3tcU6mB5xzEM05AwjnbIXHOazJMVDIseScSQrneGQKJ0MpG2rO2XzvHAYqZQg6xwg5p6wOOmcQxzm4O60ZIecYwXVsLHLMuah7NS3fU/tflnziHuuc/IB2LqN5dyCEeH2NxzyqyilSOLNJvkNW2KfMTeEzLZ2emn890AtIkUobhAoqDDmoLRRKeGCkhSk8D9QKQxpqfeBmsumZQRHHX/938h8AAP//AAAA//9UUNFKxEAM/JUlgm/aFpGD2hbOg4M7ER/UD9ijaRvsbZZsior477bblsM8TDIJTCYp7KC8p15RjGBTwjbL3zIwX5IPVJfwky5xM+ZsgvQC6+wXkqrwtsVnKy25YHpstIT0dgNGqO3WWtnH7j2YE6vyeWUd2hplYndgGubRzUwW3VfUwRsWQqdWiV0JnkXFko4b8smoHOpssjFL7aNGVXBdL+W1PfuHXcTsKibY2Z5OQhDZJuLTfNJxcOTHjzwOgRyGYN4DmhfXfxfJRbFI/u9KPlk+Qoeo1R8AAAD//wMAUEsDBBQABgAIAAAAIQB46LpVTwIAAD0EAAAYAAAAeGwvd29ya3NoZWV0cy9zaGVldDIueG1snNJfa4MwEADw98G+Q8i7jVp1VbSlXVe2tzG2vafxrKEmkST9x9h3X7S0HfSlFBTUJL+78y6f7EWDtqANV7LAwcDHCCRTJZerAn99LrwRRsZSWdJGSSjwAQyejB8f8p3Sa1MDWOQEaQpcW9tmhBhWg6BmoFqQbqVSWlDrXvWKmFYDLftDoiGh7ydEUC7xUcj0LYaqKs5grthGgLRHRENDrcvf1Lw1J02wWzhB9XrTekyJ1hFL3nB76FGMBMveVlJpumxc3fsgogzttbtCdw9PYfrvV5EEZ1oZVdmBk8kx5+vyU5ISys7Sdf03MUFENGx518ALFd6XUhCfrfCCDe/EkjPW/S6dbXhZ4J/oKUmDZz/20iCZedE8evHSJJ56wWIRjmZxlEaz+BeP85K7DndVIQ1VgacBJuO8H55vDjvz7xl1s7hUat0tvLkYfreVXO1d9LP4rlEJFd009kPtXoGvausGP3YZdi3OysMcDHOz5ZhBGDvpDwAA//8AAAD//7IpzkhNLXFJLEnUtwMAAAD//wAAAP//VI3RasMwDEV/xajQx7lljEDr5GGFQTdGH8o+wMFqIpZaRlYY+/sl7kKpHq7uFeJcl3yHn146itkMeNEaNk8VGKGuX7xyKtcXMC2r8nVJPfqAMqdnMBdmXYJt3Mw9o47JsBBG9Uoca0gsKp50athRqEGOYQvT+w31VhiN4xD+7dpf0/5QdLsqCw5+oFYISqqKfmzKvI+REop5HTNFzNl8ZTSnOPw6eyc6+9hlf1i+c4+ozR8AAAD//wMAUEsDBBQABgAIAAAAIQDBFxC+TgcAAMYgAAATAAAAeGwvdGhlbWUvdGhlbWUxLnhtbOxZzYsbNxS/F/o/DHN3/DXjjyXe4M9sk90kZJ2UHLW27FFWMzKSvBsTAiU59VIopKWXQm89lNJAAw299I8JJLTpH9EnzdgjreUkm2xKWnYNi0f+vaen955+evN08dK9mHpHmAvCkpZfvlDyPZyM2Jgk05Z/azgoNHxPSJSMEWUJbvkLLPxL259+chFtyQjH2AP5RGyhlh9JOdsqFsUIhpG4wGY4gd8mjMdIwiOfFsccHYPemBYrpVKtGCOS+F6CYlB7fTIhI+wNlUp/e6m8T+ExkUINjCjfV6qxJaGx48OyQoiF6FLuHSHa8mGeMTse4nvS9ygSEn5o+SX95xe3LxbRViZE5QZZQ26g/zK5TGB8WNFz8unBatIgCINae6VfA6hcx/Xr/Vq/ttKnAWg0gpWmttg665VukGENUPrVobtX71XLFt7QX12zuR2qj4XXoFR/sIYfDLrgRQuvQSk+XMOHnWanZ+vXoBRfW8PXS+1eULf0a1BESXK4hi6FtWp3udoVZMLojhPeDINBvZIpz1GQDavsUlNMWCI35VqM7jI+AIACUiRJ4snFDE/QCLK4iyg54MTbJdMIEm+GEiZguFQpDUpV+K8+gf6mI4q2MDKklV1giVgbUvZ4YsTJTLb8K6DVNyAvnj17/vDp84e/PX/06PnDX7K5tSpLbgclU1Pu1Y9f//39F95fv/7w6vE36dQn8cLEv/z5y5e///E69bDi3BUvvn3y8umTF9999edPjx3a2xwdmPAhibHwruFj7yaLYYEO+/EBP53EMELEkkAR6Hao7svIAl5bIOrCdbDtwtscWMYFvDy/a9m6H/G5JI6Zr0axBdxjjHYYdzrgqprL8PBwnkzdk/O5ibuJ0JFr7i5KrAD35zOgV+JS2Y2wZeYNihKJpjjB0lO/sUOMHau7Q4jl1z0y4kywifTuEK+DiNMlQ3JgJVIutENiiMvCZSCE2vLN3m2vw6hr1T18ZCNhWyDqMH6IqeXGy2guUexSOUQxNR2+i2TkMnJ/wUcmri8kRHqKKfP6YyyES+Y6h/UaQb8KDOMO+x5dxDaSS3Lo0rmLGDORPXbYjVA8c9pMksjEfiYOIUWRd4NJF3yP2TtEPUMcULIx3LcJtsL9ZiK4BeRqmpQniPplzh2xvIyZvR8XdIKwi2XaPLbYtc2JMzs686mV2rsYU3SMxhh7tz5zWNBhM8vnudFXImCVHexKrCvIzlX1nGABZZKqa9YpcpcIK2X38ZRtsGdvcYJ4FiiJEd+k+RpE3UpdOOWcVHqdjg5N4DUC5R/ki9Mp1wXoMJK7v0nrjQhZZ5d6Fu58XXArfm+zx2Bf3j3tvgQZfGoZIPa39s0QUWuCPGGGCAoMF92CiBX+XESdq1ps7pSb2Js2DwMURla9E5PkjcXPibIn/HfKHncBcwYFj1vx+5Q6myhl50SBswn3Hyxremie3MBwkqxz1nlVc17V+P/7qmbTXj6vZc5rmfNaxvX29UFqmbx8gcom7/Lonk+8seUzIZTuywXFu0J3fQS80YwHMKjbUbonuWoBziL4mjWYLNyUIy3jcSY/JzLaj9AMWkNl3cCcikz1VHgzJqBjpId1KxWf0K37TvN4j43TTme5rLqaqQsFkvl4KVyNQ5dKpuhaPe/erdTrfuhUd1mXBijZ0xhhTGYbUXUYUV8OQhReZ4Re2ZlY0XRY0VDql6FaRnHlCjBtFRV45fbgRb3lh0HaQYZmHJTnYxWntJm8jK4KzplGepMzqZkBUGIvMyCPdFPZunF5anVpqr1FpC0jjHSzjTDSMIIX4Sw7zZb7Wca6mYfUMk+5YrkbcjPqjQ8Ra0UiJ7iBJiZT0MQ7bvm1agi3KiM0a/kT6BjD13gGuSPUWxeiU7h2GUmebvh3YZYZF7KHRJQ6XJNOygYxkZh7lMQtXy1/lQ000RyibStXgBA+WuOaQCsfm3EQdDvIeDLBI2mG3RhRnk4fgeFTrnD+qsXfHawk2RzCvR+Nj70DOuc3EaRYWC8rB46JgIuDcurNMYGbsBWR5fl34mDKaNe8itI5lI4jOotQdqKYZJ7CNYmuzNFPKx8YT9mawaHrLjyYqgP2vU/dNx/VynMGaeZnpsUq6tR0k+mHO+QNq/JD1LIqpW79Ti1yrmsuuQ4S1XlKvOHUfYsDwTAtn8wyTVm8TsOKs7NR27QzLAgMT9Q2+G11Rjg98a4nP8idzFp1QCzrSp34+srcvNVmB3eBPHpwfzinUuhQQm+XIyj60hvIlDZgi9yTWY0I37w5Jy3/filsB91K2C2UGmG/EFSDUqERtquFdhhWy/2wXOp1Kg/gYJFRXA7T6/oBXGHQRXZpr8fXLu7j5S3NhRGLi0xfzBe14frivlzZfHHvESCd+7XKoFltdmqFZrU9KAS9TqPQ7NY6hV6tW+8Net2w0Rw88L0jDQ7a1W5Q6zcKtXK3WwhqJWV+o1moB5VKO6i3G/2g/SArY2DlKX1kvgD3aru2/wEAAP//AwBQSwMEFAAGAAgAAAAhAA/LtMF8CwAA1FMAAA0AAAB4bC9zdHlsZXMueG1szFxbb9vIFX4v0P9AKN1iNygjXoak6LWckhTZBkjTReOifVhgQUuUzYYXlaISeYv97z0zvM1IlDSyaY1jxBYpzjnfuc6ZC+f6/TZNpK9RsY7zbDpS3ykjKcrm+SLO7qejf94G8mQkrcswW4RJnkXT0WO0Hr2/+f3vrtflYxJ9foiiUgIS2Xo6eijL1dV4vJ4/RGm4fpevogy+WeZFGpZwWdyP16siChdr3ChNxpqimOM0jLNRReEqnfMQScPiy2Ylz/N0FZbxXZzE5SOhNZLS+dWH+ywvwrsEoG5VFM6lrWoWmrQtGibk7h6fNJ4X+Tpflu+A7jhfLuN5tA/XHtvjcN5RAspPo6QaY0VjZN8WT6SExkX0NcbmG91cZ5s0SMu1NM83WTkdae0tqfrmwwJsbKKRVFnFyxegp1++/+N/N3n54x+qP2+lN39680Z5pyi//PDj3nc/f998+3PPtxUFufrz/j0h8OdffhiNG2gMDmMXxw7rt9IOt7fSEQbjWvib62WeUToAaYnNr75k+bcswN+BEkAz+LGb6/Wv0tcwgTsqRjnPk7yQSvBgUAy5k4VpVD3hhUl8V8T4sWWYxsljdVvDN4jT18+lMbggvjmuOAjiM9mTRyd22JVH+hjfP5SnpQr/0yPVHZa90aDByfGJGmR4EVkYa/VLNwSvfc8YlFefBxb3d9NREECGVBWl12wDumHNzPYU4HcxZoZ1Mcn0QA+sQSVjfJHyj1qVmKEeDKnKEwwDx5pdTJ3DMzskXZ2ILxUBOOCG1WJM58d9P7EC/HMJP7lwl/YiFtvv0XbYOEUcJr2JsbcrpujVcdvnAGcSVZ+HqSpDuqLjNk6jtfQp+ib9I0/DbFc44lKnhOu3/bly7Xuv7ZnGUOFCZFhDZRUnSVu/6iYu0+DOzTXU+mVUZAFcSPXn28cVFGkZDEuwVsbVcyeevi/CR1UjhQpfg3WexAuM4t4jpWHtKJ7pB55P+FLIeFEcIBoEnvUCRH3X9oZH6tn20ES1AH4GJuoY+Gdw8cFUg+m0Dk80FMiWnlTGeBSovLNs256o5mQysZGuIkSUfFd7dJwtom0EA0NzMDXtIzAAga1PbFMDIAqaEFYXRaADAMswJoZqawj+k3T28giG1imMmQVblUIgyKoUAkFWJWOJ8QCZv44UU7hVKQSCrEohEGRVa+AMbAm3KoVAkFUpBIKsSqr8AWMVJsQFZ2AKgSCrUggEWXWw4rPOwLZwq1IIBFmVQiDIqgeKsqFLYzJE36/9BmOzN5twDi8yLoWR8F1eLGBdrlnLgYmB5t7NdRItSxjzFnjmHv6W+Qp+3+VlCYtXN9eLOLzPszDBo+KKCldLWOeDJb3pqHyI51+AGTO7XummYvFSHNq8ivB4BVlIsZChmdWQcCDWabSIN+m+dC3vXs8HNWLdnhac0iFej6tUWLtDN9k3xvarzcfZgpiaWJqzAfhE4xKcLYaQsZv45pWRasEnI9WAU0aqxUBetMg3sMS8a+AgmChKNSN3ri77CR7X5sk2+/o82aRHoyfbDOE3roZ/SP3NGRtUCz6/oRpw+g3V4ml+0xt4zHTPaX0zjz9L1d1ET5XMTyl69/kTat59/JSSd58/JlvdjUGvOI+S5DPup/69bLtGQ4dI3C6pvQWwiwWvuuPtDvgjTBrXH6vusLqA7pFppMIEWNtMPdhMCler5BHvIyDEqyvg0F25pOfurp0kvs/SiG7wU5GX0bwkm25INXIIvnYCx6dNehcVAdlkcwk8+gE8oC+iCR48z9EH3rvSZ07Q00X4Yxfp4w96Ecof3Jyb/5D+CbM2jT7ABHR4HcMzJAKYYWgQgBFEIIDRcIMA3FMEAhi5NQjAQTsEAOeIVzwnDmGZs2UJPtCxBP4vxZLKyQzLF5TyUPoF/hcJd5XKt+DonZrh4qXUfCjFikoxlNkh0joVwMUxGzAd9LNc/VDOF6YQKukChk4jkAYu45VUzgWeIjIe5RQ4FYmGAHBEQxDU/aqUL+BsJVoNgnpg2h+ZLvh4khiyEKIhMP3jBSHQzsD0V2LUwPQXYiCIz4+aqPxIOQNstxCdGTRRCZIaIYnPj5qo/Eg7g/gEqYlKkJQziM+Pmqj8SDuD+ASpiy8gdfH5UReUH3H3xD1DS15dqyZ2cT4/0O7CU7SHgexOMVPz0ho1sn8FOVl9DcOHV1C1vopi6TX00a+ga3jpjDSmF5OqpSVqVQmmFZ6yqiRtl/3LSxxrUngmo2le5bBmfEhWlc9d4jpGriJPTYvh7SXwgmu1UCU95EX8K8yg4Tdj57ByFVUvtG6Xh1WiUtPg+m6hy0pVLZ41S0b4nQ78su0LQKLmxfXdoOKBdHSJEI9pejugTu+tiGQysl/gzsQvbJPdqluQTWg3YSDtu+sxP+GKDMYnLyO/9K0IV7fRFlaHq60gezGjUkXModpnXxfswmrjSkM4D0mDkPioxXV2ab1NkxJ+L206+q7JUgjCfFtcbWJYZ/8ffnUS/5Phv4F/kU/Vr+a73/AOj11ymvKdJEvOHOcZkKtyS9IPb+IEttridIpL9PlmDXuw3Opmrd1jtKCiqWnh0KdoweW5tLCgJK+TIRNFC7z5XFr4QAJCi/RvHS14LedsWuDTNS1cO1G04PJcXNCkooXweKijZXDqHvXaEQ9sKH3BJQ8umlZnRxzBFC24PJdWZ0c87KRoweW5tDo74t6O0j1cnkursyOeFaBoweW5tDo7sro3OHVv9toRjw8ofXH6KkNLai2JQKw2caiWoTu+acm677syUpWJPHFcRTZNa+a7M0txFb83cdDEOydhAwFeIeVSIEOLAgrx3QLVZo5h+DNT1q3Ak5HumzLAc2TT9x17NkHIso2TQFsPJLMOlKU5o4wGqlNAwYFboF4Q+CiwbNm3kSUjwwGgpjGTvcAwHGviOQHSTgLt3JtNn4gzfdJAUQcUZ5MOqIE83zd9sHrgyEgxVNlxA1+2XGTOLCNQJz46CbSLHTxmoTTKmZtpoAYFFPdFTeemusi3TMOWHRsMjtwJfEI+kj3L8QPFdBzFNE8C7QKTDSbjCcFkUkCx77dALVuZ+eZMhpiCYEKeI7uWr8rqTEco0CzNm7m9QHf6X5XNjxpnfqyodMHIZkY8LODJZhWVLlezOVHnzIkVlc6N2WyIOLNhRaXzMbYPQpx9UEWldQCd1S7i1K4bLpo+mnV1jVMlcITRfJPAsVlwWlTd2ZOJny5o8Kwaj4m8h2j+RfJgN2ZLiE0TeH6Li1CepmFLg+1k+qpFb1MUcELZo4SB1/UKtjFPGarCGSn4X28A+NtVEmZhmRePEq7fW0ys98FmUy65/pLnrbFYCnjtgUczf4Wz0uAYNqmtidlgUs8k06mLUTJevz8HTRuVbCDgkdY5ZNqwZLtrnHZ4yHzIVpvOQmy9iocOPDQ+xtmXaMH6MKtivFjGQ+lTtCkLOHmk9sed1MmpmZoI1aXjdNG6tW7qjjqzZc/XIa/jSgmKEUP2FOjeNdcyA7+/SPqEtwa30FhNkZ3Hu0OoqgEdXjj58YSXBq8/HAyvT7DhuQ2pHaNzGuzvm5K2OpvP8WwMj61u4xLe3GhGc2wocCbz2xzmxVpZdjIxJ41/hUWGg5vJNDshdUCibtoShuyLbbcPnpizxIcfkh3y7SAe1LuIluEmKW/bL6ej7vPfyDtRkB/qp36Kv+YlITEddZ/JMXGQdGAED9nx4xpeYIK/0qaIYfDvu5Y98wNNnijuBErjyJBtw4UqBHnubBbYiqZ4v4HK8EmRV3Ac3zMOYCQnRsI0pIqu1gkc01jUwtbgP3f3piPqooJP3jAB2DR2WzMVx1AVOdAVVUZmCOMPUzfkwFC1mYlc3wigrG+xG088qFEZq2p15CMGb1yVcLJPEmeNrRoL0XfBSHB5RIhxY4lxdxznzf8BAAD//wMAUEsDBBQABgAIAAAAIQCK3HpNqAIAAGQHAAAUAAAAeGwvc2hhcmVkU3RyaW5ncy54bWyUVV1v2jAUfZ+0/2DladNE88VHqSBVFlpCCxmDsGl7qdzkAm4TO7Wdqvz7OWyTwKZS+xjfe3zOvfdcZ3D5UhboGbggjA4t98yxENCM5YRuhtYqvW6dW0hITHNcMApDawfCugw+fhgIIZHCUjG0tlJWF7Ytsi2UWJyxCqiKrBkvsVSffGOLigPOxRZAloXtOU7XLjGhFspYTeXQ6nsWqil5qiH6e9DxrWAgSDCQwY0KVMDRFZVE7lDEchjYMhjYTXif0ui4EBXOlD5FJIA/gxWMp/tcpCUHYdlQtNColjt7Ml6mekJUc66fjXir7dgRb3UcPTShz2iEpS4qmBIKaAQi46SSqrsGDVMdjIBKMNjmnD1AJtFkdIKMkQxQUpf3Ji5WTVadOiBFkiFR35dEop8Aj8UOXZPCkBoKQTa0VFrQAtbAlQOMnBS/nOp9MEkWusixp58k0QL9AJozriZuzC/4SooCcnTNWakj/4VSZpAs00nyJiPEy2TPajghZRIX6BPeG+KLxC+fdZKm6DWBIj9V0LdMMjUE5Dle2/DL1TSe6IezMA4X4TJOF6EecnphGEZJz233p+5v47px/YA5ljrKax+izvVwjPkOU2xwdQ9RXT18iznFEj8aOK9/iPtlDMQ16sVbpVpsJTfvOqrX8EuKS1KgBOe1fqfvH4roGwODAtPNiaL9o6J9HbeSEnM052p/xNZo2FHhxgMww+qBqo0ldo+GYyhVK3E3oRnJm61bA6BVVTCc3x24yhhNs0bpfIacdmO6LlJPz/41iJazxO16r+Z/N/JXc7fn+KdE7QlMQLrs93qdU2vQCDqRn3ie3zaM/L8CEzCLPc/t9t7OcBu6fsc3TP86Q7xwHc99B8P4xnX89xQ9mrpO3znYA1v9JYM/AAAA//8DAFBLAwQUAAYACAAAACEAO20yS8EAAABCAQAAIwAAAHhsL3dvcmtzaGVldHMvX3JlbHMvc2hlZXQxLnhtbC5yZWxzhI/BisIwFEX3A/5DeHuT1oUMQ1M3IrhV5wNi+toG25eQ9xT9e7McZcDl5XDP5Tab+zypG2YOkSzUugKF5GMXaLDwe9otv0GxOOrcFAktPJBh0y6+mgNOTkqJx5BYFQuxhVEk/RjDfsTZsY4JqZA+5tlJiXkwyfmLG9Csqmpt8l8HtC9Ote8s5H1Xgzo9Uln+7I59Hzxuo7/OSPLPhEk5kGA+okg5yEXt8oBiQet39p5rfQ4Epm3My/P2CQAA//8DAFBLAwQUAAYACAAAACEAE8QsE8IAAABCAQAAIwAAAHhsL3dvcmtzaGVldHMvX3JlbHMvc2hlZXQyLnhtbC5yZWxzhI/BasMwEETvhfyD2HskO4dQiiVfSiHXJv0ARV7bovZKaLcl+fvo2IRCjsNj3jBdf1kX9YuFYyILrW5AIYU0RJosfJ0+tq+gWDwNfkmEFq7I0LvNS/eJi5da4jlmVtVCbGEWyW/GcJhx9axTRqpkTGX1UmOZTPbh209odk2zN+WvA9ydUx0GC+UwtKBO11yXn7vTOMaA7yn8rEjyz4TJJZJgOaJIPchV7cuEYkHrR/aYd/ocCYzrzN1zdwMAAP//AwBQSwMEFAAGAAgAAAAhAM/gR7TCAQAALBUAACcAAAB4bC9wcmludGVyU2V0dGluZ3MvcHJpbnRlclNldHRpbmdzMS5iaW7sVM1K3FAYPTOx7eimDhTcdFGkK3HoDJOp3VWZpHZK0oQkM7hxMXRSCIzJkEREpYL4Gj5Ily5d9gG6diHFB3Cj56Yz2JahjOBG+O7lu9/PPTk395B8NiJ8QYoEGe0rcryCyzxCXMQ5q6pi4AOmjdKc9vQn3BfamxLUvFxIKgP659gql+m3yhpXCyHZcq7pVJb7FUtjuPJlmvI3HJsdX/+Tyeh87i7jHK+11er77cOj/53ypNicL7ge4BWF4hEqMPmuZnn1c4J8O/iksIv4jkPU8Q46/5I6Glw3UIOJt2iyVqMZWOOsEdNk3WRUZ64zb9C3mTXRKrJvZPRM37AsdOMoDTMVuf1RmPrRQQjLDALTg5NGYZz38yiJ4Tpe4G10Anhhlgx3ixpDZ6SiBtrJMEntZBD+jv6+3WoV6OmGPbn76cJo+SUhv2ga7brkVPSLPfvk6tnHpbPW8Q/WrPEeKndcCqvylbFX+Tqtp/JF8P4J+8wudtgDVGfpst+obuCizyjDHvdTDAj+F+lwL54R2ybHPkbk9/mEOk91spw1GaKAKCAKiAKigCggCogCooAoIAqIAqKAKCAKzKLALQAAAP//AwBQSwMEFAAGAAgAAAAhAM/gR7TCAQAALBUAACcAAAB4bC9wcmludGVyU2V0dGluZ3MvcHJpbnRlclNldHRpbmdzMi5iaW7sVM1K3FAYPTOx7eimDhTcdFGkK3HoDJOp3VWZpHZK0oQkM7hxMXRSCIzJkEREpYL4Gj5Ily5d9gG6diHFB3Cj56Yz2JahjOBG+O7lu9/PPTk395B8NiJ8QYoEGe0rcryCyzxCXMQ5q6pi4AOmjdKc9vQn3BfamxLUvFxIKgP659gql+m3yhpXCyHZcq7pVJb7FUtjuPJlmvI3HJsdX/+Tyeh87i7jHK+11er77cOj/53ypNicL7ge4BWF4hEqMPmuZnn1c4J8O/iksIv4jkPU8Q46/5I6Glw3UIOJt2iyVqMZWOOsEdNk3WRUZ64zb9C3mTXRKrJvZPRM37AsdOMoDTMVuf1RmPrRQQjLDALTg5NGYZz38yiJ4Tpe4G10Anhhlgx3ixpDZ6SiBtrJMEntZBD+jv6+3WoV6OmGPbn76cJo+SUhv2ga7brkVPSLPfvk6tnHpbPW8Q/WrPEeKndcCqvylbFX+Tqtp/JF8P4J+8wudtgDVGfpst+obuCizyjDHvdTDAj+F+lwL54R2ybHPkbk9/mEOk91spw1GaKAKCAKiAKigCggCogCooAoIAqIAqKAKCAKzKLALQAAAP//AwBQSwMEFAAGAAgAAAAhAH6BLwZfAQAAeQIAABEACAFkb2NQcm9wcy9jb3JlLnhtbCCiBAEooAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHySX0/CMBTF3038DkvfR9uhRJsxgn94QRITMRrfanuB6to1bRH49nYb4IjGx/ac++s5N81HW10mX+C8qswQ0R5BCRhRSWWWQ/Q8n6RXKPGBG8nLysAQ7cCjUXF+lgvLROXg0VUWXFDgk0gyngk7RKsQLMPYixVo7nvRYaK4qJzmIR7dElsuPvkScEbIAGsIXPLAcQ1M7ZGI9kgpjki7dmUDkAJDCRpM8Jj2KP7xBnDa/znQKB2nVmFnY6d93C5bilY8urdeHY2bzaa36TcxYn6KX2cPT03VVJl6VwJQkUvBhAMeKleMP/guma41d8m0shaCynFHrldZch9mcesLBfJmV4yd4sZAMouD917wd1VGR45/O+M7Ta32MZBJDMraWgflpX97N5+gIiP0OiUXKSVzmrFLwvqDtzrIyXwdvL3Q+zj/ErNBxDVEykjGaNYhHgBFk/v0sxTfAAAA//8DAFBLAwQUAAYACAAAACEAzYvFXYQBAAAWAwAAEAAIAWRvY1Byb3BzL2FwcC54bWwgogQBKKAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACckk9P4zAQxe9IfIfId+q0u0KocoxWZRGHha1ogbPXmTQWrh15hqjdT88kESWFAxK3+fP08svzqMvd1mctJHQxFGI6yUUGwcbShU0hHtbXZxciQzKhND4GKMQeUFzq0xO1TLGBRA4wY4uAhaiJmrmUaGvYGpzwOvCmimlriNu0kbGqnIWraF+2EEjO8vxcwo4glFCeNQdDMTjOW/quaRltx4eP633DwFr9ahrvrCH+S33rbIoYK8p+7yx4JcdLxXQrsC/J0V7nSo5btbLGw4KNdWU8gpLvA3UDpgttaVxCrVqat2Appgzdf45tJrJ/BqHDKURrkjOBGKuTDU1f+wYp6aeYnrEGIFSSBcOwL8face1+6lkv4OJY2BkMILw4Rlw78oB/q6VJ9BVxzzDwDjh3i/sx3AFz1XFPP3H3UTDBh2/+ceEZH5p1vDIEb5keD9WqNglKfoZD5oeBuuE4k+9MFrUJGyjfNJ8X3QU8Dmeup+eT/EfOjzuaKfl+0PoVAAD//wMAUEsDBBQABgAIAAAAIQCJoo+NugEAAO0FAAATAAgBZG9jUHJvcHMvY3VzdG9tLnhtbCCiBAEooAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAALSUUWvbMBSF3wf7D0bvqmVJtmVjpzRxCoV1FJLtYS9BluTGYEvGUrKFsf9eZVsbMhgMDz+KK46+czk6xe23vguOarSt0SWIbhAIlBZGtvq5BJ+295CBwDquJe+MViU4KQtuF+/fFU+jGdToWmUDL6FtCfbODXkYWrFXPbc3fqz9pDFjz50/js+haZpWqMqIQ6+0CzFCSSgO1pkeDm9y4JdefnRTJaURZzr7eXsaPO6i+C1+CpretbIE36t4VVUxiiFeZysYoWgJM5KlEDGE8BKv7rO79Q8QDOfLGASa99764+bhafeB16rboYSQmjEGuUISUkIjyNO4gShRiEYklWkc79aa152S/v2jy7vhq3Xjwo0HVYSXcxG+sv0nJZlKuVGu4k5dUWKECUQEYrSNWE6SHNEvs1DTqdSPyu3N9Wo3P1M6yllA46mgH310rnb7L9GZxUIyOSGtUw/Xu64VT1lNBKRC1tAnHkEWUw4jGWUEE6zihs5iIp1q4k4433B/2GCEcMoiBqMGI2/D/2VeUwFjIlJcE97UEs1iw5fqtE5ZGe18dS5bd+61S4/gv2GGl55evAAAAP//AwBQSwECLQAUAAYACAAAACEA1XYZ9X0BAAAJBgAAEwAAAAAAAAAAAAAAAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQATXr5lAgEAAN8CAAALAAAAAAAAAAAAAAAAALYDAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQA4UWf5PQMAAIgHAAAPAAAAAAAAAAAAAAAAAOkGAAB4bC93b3JrYm9vay54bWxQSwECLQAUAAYACAAAACEASqmmYfoAAABHAwAAGgAAAAAAAAAAAAAAAABTCgAAeGwvX3JlbHMvd29ya2Jvb2sueG1sLnJlbHNQSwECLQAUAAYACAAAACEA4Xu7U4wHAAAwIwAAGAAAAAAAAAAAAAAAAACNDAAAeGwvd29ya3NoZWV0cy9zaGVldDEueG1sUEsBAi0AFAAGAAgAAAAhAHjoulVPAgAAPQQAABgAAAAAAAAAAAAAAAAATxQAAHhsL3dvcmtzaGVldHMvc2hlZXQyLnhtbFBLAQItABQABgAIAAAAIQDBFxC+TgcAAMYgAAATAAAAAAAAAAAAAAAAANQWAAB4bC90aGVtZS90aGVtZTEueG1sUEsBAi0AFAAGAAgAAAAhAA/LtMF8CwAA1FMAAA0AAAAAAAAAAAAAAAAAUx4AAHhsL3N0eWxlcy54bWxQSwECLQAUAAYACAAAACEAitx6TagCAABkBwAAFAAAAAAAAAAAAAAAAAD6KQAAeGwvc2hhcmVkU3RyaW5ncy54bWxQSwECLQAUAAYACAAAACEAO20yS8EAAABCAQAAIwAAAAAAAAAAAAAAAADULAAAeGwvd29ya3NoZWV0cy9fcmVscy9zaGVldDEueG1sLnJlbHNQSwECLQAUAAYACAAAACEAE8QsE8IAAABCAQAAIwAAAAAAAAAAAAAAAADWLQAAeGwvd29ya3NoZWV0cy9fcmVscy9zaGVldDIueG1sLnJlbHNQSwECLQAUAAYACAAAACEAz+BHtMIBAAAsFQAAJwAAAAAAAAAAAAAAAADZLgAAeGwvcHJpbnRlclNldHRpbmdzL3ByaW50ZXJTZXR0aW5nczEuYmluUEsBAi0AFAAGAAgAAAAhAM/gR7TCAQAALBUAACcAAAAAAAAAAAAAAAAA4DAAAHhsL3ByaW50ZXJTZXR0aW5ncy9wcmludGVyU2V0dGluZ3MyLmJpblBLAQItABQABgAIAAAAIQB+gS8GXwEAAHkCAAARAAAAAAAAAAAAAAAAAOcyAABkb2NQcm9wcy9jb3JlLnhtbFBLAQItABQABgAIAAAAIQDNi8VdhAEAABYDAAAQAAAAAAAAAAAAAAAAAH01AABkb2NQcm9wcy9hcHAueG1sUEsBAi0AFAAGAAgAAAAhAImij426AQAA7QUAABMAAAAAAAAAAAAAAAAANzgAAGRvY1Byb3BzL2N1c3RvbS54bWxQSwUGAAAAABAAEABTBAAAKjsAAAAA"
    india_pdfs = st.file_uploader("Upload matching invoice PDFs", type=["pdf"], accept_multiple_files=True, key="india_pdfs")
    if india_pdfs:
        if st.button("Generate India CSV", type="primary", use_container_width=True):
            try:
                with st.spinner("Matching invoice numbers and supplier GSTINs..."):
                    template_bytes = base64.b64decode(_DEFAULT_TEMPLATE_B64)
                    india_frame, india_results, india_unmatched, india_duplicates = _india_make_csv(
                        template_bytes, india_pdfs, selected_quarter, int(selected_year),
                        selected_fee_type, selected_upload_month, int(selected_year)
                    )
                st.session_state["india_csv_bytes"] = india_frame.to_csv(index=False, date_format="%m/%d/%Y").encode("utf-8-sig")
                st.session_state["india_csv_name"] = f"NCR Invoice AP Upload file {datetime.now().strftime('%m.%d.%Y')}.csv"
                st.session_state["india_results"] = india_results
                st.session_state["india_unmatched"] = india_unmatched
                st.session_state["india_duplicates"] = india_duplicates
                st.success(f"CSV prepared with {len(india_results)} matched invoice record(s).")
            except Exception:
                st.error("The built-in NCR template or invoice PDFs could not be processed. Check the uploaded PDFs.")
        if st.session_state.get("india_results") is not None:
            st.subheader("Matched invoice review")
            review = pd.DataFrame(st.session_state["india_results"])
            st.dataframe(review, use_container_width=True, hide_index=True)

            # Totals for the three requested AP upload amount columns.
            # Calculate from the generated workbook frame and ignore blank/non-numeric cells.
            _india_totals_frame = pd.read_csv(
                BytesIO(st.session_state["india_csv_bytes"]),
                encoding="utf-8-sig",
            )
            _india_total_columns = [
                "Amount- Duty/IGST",
                "Tax field",
                "Total (amount+tax)",
            ]
            _india_totals_rows = []
            for _column in _india_total_columns:
                _numeric_values = pd.to_numeric(
                    _india_totals_frame[_column], errors="coerce"
                ).dropna()
                _india_totals_rows.append({
                    "Amount column": _column,
                    "Total amount": round(float(_numeric_values.sum()), 2),
                })
            st.subheader("Total Amount Summary")
            _india_totals_display = pd.DataFrame(_india_totals_rows)
            _india_totals_display["Total amount"] = _india_totals_display["Total amount"].map(
                lambda value: f"{value:,.2f}"
            )
            st.dataframe(
                _india_totals_display,
                use_container_width=True,
                hide_index=True,
            )

            if st.session_state.get("india_unmatched"):
                st.warning("Some template rows did not match both invoice number and supplier GSTIN. Those rows were not filled: " + ", ".join(st.session_state["india_unmatched"]))
            if st.session_state.get("india_duplicates"):
                st.warning("Duplicate invoice PDFs were uploaded for: " + ", ".join(st.session_state["india_duplicates"]))
            st.download_button(
                "Download NCR Invoice AP Upload CSV",
                data=st.session_state["india_csv_bytes"],
                file_name=st.session_state["india_csv_name"],
                mime="text/csv",
                use_container_width=True,
            )
    else:
        st.info("Upload the corresponding invoice PDFs to generate the India CSV using the built-in NCR template.")
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
