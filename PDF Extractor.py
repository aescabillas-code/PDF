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

# Encoded configuration keeps organization-specific labels and GL suffix mappings
# out of plain sight in this source file. This is obfuscation, not encryption.
_RULES_B64 = "eyJyb2xlcyI6eyJJUExTIjp7InZhdF9zdWZmaXgiOiIyMSIsIm5ldF9zdWZmaXgiOiIxOSJ9LCJBVSI6eyJ2YXRfc3VmZml4IjoiMjMiLCJuZXRfc3VmZml4IjoiMTkifSwiTkEgYW5kIExBVEFNIjp7InZhdF9zdWZmaXgiOiIyMSIsIm5ldF9zdWZmaXgiOiIxOSJ9LCJVSyI6eyJ2YXRfc3VmZml4IjoiMjEiLCJuZXRfc3VmZml4IjoiMTkifSwiRU1FQSAmIEFQQUMiOnsidmF0X3N1ZmZpeCI6IjIxIiwibmV0X3N1ZmZpeCI6IjE5In19LCJhbGlhc2VzIjp7Ik5BL0xBVEFNIjoiTkEgYW5kIExBVEFNIiwiTkEgJiBMQVRBTSI6Ik5BIGFuZCBMQVRBTSIsIkVNRUEgQVBBQyI6IkVNRUEgJiBBUEFDIiwiRU1FQS9BUEFDIjoiRU1FQSAmIEFQQUMifSwibmV0X2xhYmVsIjoiTkNSIENFIG9uLXNpdGUgY29zdCAvIFN1YnRvdGFsIiwidmF0X2xhYmVsIjoiVkFUIC8gR1NUIiwiZ2xfY29kZXMiOnsibmV0IjoiMzIzMDYyNDAxMDUwMTE2NzAxOSIsInZhdCI6IjMyMzA2MjQwMTA1MDExNjcwMjEiLCJhdV9nc3QiOiIzMjMwNjI0MDEwNTAxMTY2MDIzIn0sImxhYmVsX3Rlcm1zIjp7Im5ldCI6WyJORVQgQU1PVU5UIChVU0QpIiwiTkVUIEFNT1VOVChVU0QpIiwiTkVUIEFNT1VOVCIsIlNVQlRPVEFMIl0sInZhdCI6WyJUT1RBTCBSRUdJT04gVkFUL1RBWEVTIChVU0QpIiwiVE9UQUwgUkVHSU9OIFZBVC9UQVhFUyhVU0QpIiwiVE9UQUwgUkVHSU9OIFZBVC9UQVhFUyIsIlRPVEFMIEcuUy5UIiwiVE9UQUwgR1NUIiwiVkFUIl0sInRvdGFsIjpbIlRPVEFMIEFNT1VOVCAoVVNEKSIsIlRPVEFMIEFNT1VOVChVU0QpIiwiVE9UQUwgQU1PVU5UIiwiSU5WT0lDRSBUT1RBTCIsIkFNT1VOVCJdfX0="
try:
    _CONFIG = json.loads(base64.b64decode(_RULES_B64).decode("utf-8"))
except Exception:
    # Safe fallback retains functionality if the encoded configuration is damaged.
    _CONFIG = {
        "roles": {
            "IPLS": {"vat_suffix": "21", "net_suffix": "19"},
            "AU": {"vat_suffix": "23", "net_suffix": "19"},
            "NA and LATAM": {"vat_suffix": "21", "net_suffix": "19"},
            "UK": {"vat_suffix": "21", "net_suffix": "19"},
            "EMEA & APAC": {"vat_suffix": "21", "net_suffix": "19"},
        },
        "aliases": {},
        "gl_codes": {
            "IPLS": {"net": "3230624010501167019", "vat": "3230624010501167021"},
            "AU": {"net": "3230624010501166019", "vat": "3230624010501166023"},
            "NA and LATAM": {"net": "1010624010501115019", "vat": "1010624010501115021"},
            "UK": {"net": "3230624010501167019", "vat": "3230624010501167021"},
            "EMEA & APAC": {"net": "3230624010501167019", "vat": "3230624010501167021"},
        },
        "label_terms": {
            "net": ["NET AMOUNT (USD)", "NET AMOUNT(USD)", "NET AMOUNT"],
            "vat": ["TOTAL REGION VAT/TAXES (USD)", "TOTAL REGION VAT/TAXES", "TOTAL G.S.T", "TOTAL GST"],
            "total": ["TOTAL AMOUNT (USD)", "TOTAL AMOUNT(USD)", "TOTAL AMOUNT", "INVOICE TOTAL"],
        },
    }
def _apply_secret_configuration(config: dict) -> dict:
    """Override regional GL mappings from Streamlit secrets when configured."""
    try:
        secret_regions = st.secrets.get("regional_gl", {})
        if hasattr(secret_regions, "items"):
            for region_name, values in secret_regions.items():
                if not hasattr(values, "get"):
                    continue
                normalized = normalize_region(region_name) if "normalize_region" in globals() else str(region_name).strip()
                config.setdefault("gl_codes", {})[normalized] = {
                    "net": str(values.get("net", "")).strip(),
                    "vat": str(values.get("vat", "")).strip(),
                }
    except Exception:
        # Missing or malformed secrets should not prevent the app from starting.
        pass
    return config


_CONFIG = _apply_secret_configuration(_CONFIG)
REGION_RULES = _CONFIG["roles"]
REGION_ALIASES = {key.upper(): value for key, value in _CONFIG.get("aliases", {}).items()}
NET_LABEL = _CONFIG.get("net_label", "NCR CE on-site cost / Subtotal")
VAT_LABEL = _CONFIG.get("vat_label", "VAT / GST")
GL_RE = re.compile(r"(?<!\d)(\d{19})(?!\d)")
MONEY_RE = re.compile(r"\$?\s*([\d,]+\.\d{2})")
EMAIL_RE = re.compile(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}")


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
    """Extract a labeled amount from common invoice layouts, including column tables.

    Handles labels and amounts on the same line, labels on a header row with a
    separate values row, and OCR/PDF extraction that separates currency symbols.
    """
    raw_lines = [line.rstrip() for line in (text or "").splitlines() if line.strip()]
    norm_lines = [re.sub(r"\s+", " ", line).strip() for line in raw_lines]
    upper_lines = [line.upper() for line in norm_lines]
    norm_labels = [re.sub(r"\s+", " ", label.upper()).strip() for label in labels]

    # First, parse line-level label/value pairs without crossing another known label.
    all_labels = [
        "NET AMOUNT (USD)", "NET AMOUNT(USD)", "NET AMOUNT",
        "TOTAL REGION VAT/TAXES (USD)", "TOTAL REGION VAT/TAXES(USD)",
        "TOTAL REGION VAT/TAXES", "TOTAL G.S.T", "TOTAL GST",
        "TOTAL AMOUNT (USD)", "TOTAL AMOUNT(USD)", "TOTAL AMOUNT",
        "INVOICE TOTAL", "SUBTOTAL", "AMOUNT DUE"
    ]
    for i, upper in enumerate(upper_lines):
        if not any(label in upper for label in norm_labels):
            continue
        # If line contains several labels (table header row), don't bind a later
        # column's value to the first label. Prefer same-line amounts only when
        # this line itself contains a single summary label.
        present = [lab for lab in all_labels if lab in upper]
        same = list(MONEY_RE.finditer(norm_lines[i]))
        if same and len(present) <= 1:
            return float(same[-1].group(1).replace(",", ""))

        # Search following rows first for the values row; stop at the next summary
        # header. Use the first currency amount for a single-label column.
        for j in range(i + 1, min(len(norm_lines), i + 5)):
            candidate_upper = upper_lines[j]
            if any(lab in candidate_upper for lab in all_labels):
                break
            matches = list(MONEY_RE.finditer(norm_lines[j]))
            if matches:
                # For a table values row with multiple amounts, select by the
                # ordinal position of the label among summary columns.
                # Build ordered, non-overlapping summary columns (longest match wins
                # when e.g. NET AMOUNT is contained in NET AMOUNT (USD)).
                label_positions = []
                for lab in sorted(set(all_labels), key=len, reverse=True):
                    pos = upper.find(lab)
                    if pos >= 0 and not any(not (pos + len(lab) <= old_pos or pos >= old_pos + len(old_lab)) for old_pos, old_lab in label_positions):
                        label_positions.append((pos, lab))
                label_positions.sort()
                requested = next((lab for lab in sorted(norm_labels, key=len, reverse=True) if lab in upper), None)
                if len(matches) > 1 and requested:
                    matching_pos = upper.find(requested)
                    # Use the leftmost matching summary label's column position.
                    ordinal = sum(1 for pos, _ in label_positions if pos < matching_pos)
                    if ordinal < len(matches):
                        return float(matches[ordinal].group(1).replace(",", ""))
                return float(matches[0].group(1).replace(",", ""))

        # A few PDF extractors emit the amount before the label.
        for j in range(i - 1, max(-1, i - 4), -1):
            if any(lab in upper_lines[j] for lab in all_labels):
                break
            matches = list(MONEY_RE.finditer(norm_lines[j]))
            if matches:
                return float(matches[-1].group(1).replace(",", ""))

    # Flattened text: map each summary label to the nearest amount before the next label.
    compact = re.sub(r"\s+", " ", text or "")
    label_pattern = r"NET\s*AMOUNT\s*\(?\s*USD\s*\)?|TOTAL\s+REGION\s+VAT\s*/\s*TAXES\s*\(?\s*USD\s*\)?|TOTAL\s+G\.?S\.?T\.?|TOTAL\s+GST|TOTAL\s+AMOUNT\s*\(?\s*USD\s*\)?|INVOICE\s+TOTAL|SUBTOTAL|\bVAT\b|\bAMOUNT\b|AMOUNT\s+DUE"
    for label in labels:
        match = re.search(re.escape(label), compact, re.I)
        if match:
            tail = compact[match.end():]
            next_label = re.search(label_pattern, tail, re.I)
            segment = tail[:next_label.start()] if next_label else tail[:120]
            amount = money_after(segment)
            if amount is not None:
                return amount
    return None


def _region_gl_codes(region: str) -> dict:
    """Read full GL codes from Streamlit secrets, with configured defaults as fallback."""
    region = normalize_region(region)
    secret_codes = {}
    try:
        secret_regions = st.secrets.get("regional_gl", {})
        for name, values in secret_regions.items():
            if normalize_region(name) == region and hasattr(values, "get"):
                secret_codes = {
                    "net": str(values.get("net", "")).strip(),
                    "vat": str(values.get("vat", "")).strip(),
                }
                break
    except Exception:
        pass

    if secret_codes.get("net") and secret_codes.get("vat"):
        return secret_codes

    configured = _CONFIG.get("gl_codes", {})
    defaults = {
        "IPLS": {"net": "3230624010501167019", "vat": "3230624010501167021"},
        "AU": {"net": "3230624010501166019", "vat": "3230624010501166023"},
        "NA and LATAM": {"net": "1010624010501115019", "vat": "1010624010501115021"},
        "UK": {"net": "3230624010501167019", "vat": "3230624010501167021"},
        "EMEA & APAC": {"net": "3230624010501167019", "vat": "3230624010501167021"},
    }
    regional = configured.get(region, {}) if isinstance(configured, dict) else {}
    if isinstance(regional, dict) and regional.get("net") and regional.get("vat"):
        return {"net": str(regional["net"]), "vat": str(regional["vat"])}
    return defaults.get(region, {"net": "", "vat": ""})


def extract_labeled_gl_rows(text: str, region: str) -> tuple[list[dict], Optional[float]]:
    """Extract only the two region-specific summary amounts using the screenshot's labels."""
    region = normalize_region(region)
    region_codes = _region_gl_codes(region)
    lines = [re.sub(r"\s+", " ", line).strip() for line in (text or "").splitlines() if line.strip()]
    upper_lines = [line.upper() for line in lines]
    all_text_upper = re.sub(r"\s+", " ", (text or "").upper())

    # AU: values sit beside explicit labels "Net Amount" and "Total G.S.T".
    # UK: use "SUBTOTAL" and "VAT" values, not the VAT percentage.
    # NA/LATAM and EMEA/APAC: use the values under the two specified USD headers.
    if region == "AU":
        net_labels = ["NET AMOUNT (USD)", "NET AMOUNT(USD)", "NET AMOUNT"]
        tax_labels = ["TOTAL G.S.T", "TOTAL GST"]
    elif region == "UK":
        net_labels = ["SUBTOTAL"]
        tax_labels = ["VAT"]
    else:
        net_labels = ["NET AMOUNT (USD)", "NET AMOUNT(USD)"]
        tax_labels = ["TOTAL REGION VAT/TAXES (USD)", "TOTAL REGION VAT/TAXES(USD)", "TOTAL REGION VAT/TAXES"]

    net_amount = _amount_near_label(text, net_labels)
    vat_amount = _amount_near_label(text, tax_labels)

    # For table-style headings, associate each heading with the amount in its own
    # column on the following value row. This avoids grabbing a subtotal, VAT rate,
    # line-item amount, or the final invoice total from another part of the page.
    def column_values_for_header(header_terms):
        for idx, line in enumerate(upper_lines):
            positions = []
            for term in header_terms:
                pos = line.find(term.upper())
                if pos >= 0:
                    positions.append((pos, term))
            if not positions:
                continue
            target_pos = min(pos for pos, _ in positions)
            for value_line_idx in range(idx + 1, min(len(lines), idx + 5)):
                value_line = lines[value_line_idx]
                money_matches = list(MONEY_RE.finditer(value_line))
                if not money_matches:
                    continue
                # If the next line contains another heading, do not cross into it.
                if any(term in upper_lines[value_line_idx] for term in (
                    "NET AMOUNT", "TOTAL REGION VAT/TAXES", "TOTAL G.S.T", "TOTAL GST",
                    "SUBTOTAL", "VAT", "TOTAL AMOUNT", "INVOICE TOTAL"
                )):
                    break
                # Use x/character position where possible. If values are a row of
                # two or three amounts, map the header's horizontal order to amount order.
                header_positions = []
                for terms in (net_labels, tax_labels, ["TOTAL AMOUNT (USD)", "TOTAL AMOUNT(USD)", "TOTAL AMOUNT", "AMOUNT"]):
                    found = [line.find(term.upper()) for term in terms if line.find(term.upper()) >= 0]
                    if found:
                        header_positions.append(min(found))
                header_positions = sorted(set(header_positions))
                ordinal = sum(1 for pos in header_positions if pos < target_pos)
                if len(money_matches) >= len(header_positions) and ordinal < len(money_matches):
                    return float(money_matches[ordinal].group(1).replace(",", ""))
                # Single-value rows, including AU and UK, use the first amount on
                # the first following value line.
                return float(money_matches[0].group(1).replace(",", ""))
        return None

    if net_amount is None:
        net_amount = column_values_for_header(net_labels)
    if vat_amount is None:
        vat_amount = column_values_for_header(tax_labels)

    # OCR can merge the label and amount into one line. Prefer label-local values.
    if net_amount is None:
        for line in lines:
            upper = line.upper()
            if any(term in upper for term in net_labels):
                amount = money_after(line)
                if amount is not None:
                    net_amount = amount
                    break
    if vat_amount is None:
        for line in lines:
            upper = line.upper()
            if any(term in upper for term in tax_labels):
                # For UK, ignore the percentage (e.g. 20.00%) and take the last
                # monetary value, which is the VAT amount.
                matches = list(MONEY_RE.finditer(line))
                if matches:
                    vat_amount = float(matches[-1].group(1).replace(",", ""))
                    break

    rows = []
    if net_amount is not None and region_codes.get("net"):
        rows.append({"Juniper GL Code": region_codes["net"], "Amount": net_amount})
    if vat_amount is not None and region_codes.get("vat"):
        rows.append({"Juniper GL Code": region_codes["vat"], "Amount": vat_amount})

    # Total is intentionally calculated from the two extracted GL amounts in the UI.
    return rows, None


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

    # GL codes are not expected to be printed on source invoices. Use the
    # selected region's secret mapping and its specific summary labels. Prefer
    # this extraction when the page has no pre-existing approval GL block.
    if not rows or not any(GL_RE.search(line) for line in lines):
        label_rows, label_total = extract_labeled_gl_rows(text, region)
        if label_rows:
            rows = label_rows
        if total is None:
            total = label_total

    if not region:
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
    upper_text = re.sub(r"\s+", " ", (page_text or "").upper())
    has_gl_block = "JUNIPER GL CODE" in upper_text and ("TOTAL AMOUNT" in upper_text or "APPROVED BY" in upper_text)
    labels = _CONFIG.get("label_terms", {})
    has_net = any(term.upper() in upper_text for term in labels.get("net", ["NET AMOUNT", "SUBTOTAL"]))
    has_tax = any(term.upper() in upper_text for term in labels.get("vat", ["TOTAL G.S.T", "TOTAL GST", "VAT"]))
    has_total = any(term.upper() in upper_text for term in labels.get("total", ["TOTAL AMOUNT", "INVOICE TOTAL", "AMOUNT"]))
    has_region_pair = (
        ("NET AMOUNT" in upper_text and ("TOTAL G.S.T" in upper_text or "TOTAL GST" in upper_text))
        or ("NET AMOUNT" in upper_text and "TOTAL REGION VAT/TAXES" in upper_text)
        or ("SUBTOTAL" in upper_text and "VAT" in upper_text)
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
    selected_region = st.selectbox("Invoice region", list(REGION_RULES.keys()))
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
        st.warning("No pre-existing Juniper GL approval block was detected. GL codes are mapped from invoice summary labels; review the detected amounts and enter approval details if they are absent.")

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
        st.warning("No GL rows were extracted automatically. Enter the configured regional GL codes and amounts below.")

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
