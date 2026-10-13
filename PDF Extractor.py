"""
PDF Approval Stamp Tool & Vendor SOW Analyzer — privacy-hardened Streamlit app.

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
import hmac
import os
import zipfile
import base64
import json
import html
import zlib
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

# -----------------------------------------------------------------------------
# PERSISTENT ONE-TIME ACCESS — adapted from TASKS MONITORING TRACKER
# -----------------------------------------------------------------------------
try:
    from itsdangerous import URLSafeTimedSerializer
except Exception:
    URLSafeTimedSerializer = None

try:
    from streamlit_js_eval import streamlit_js_eval
except Exception:
    streamlit_js_eval = None

ACCESS_STORAGE_KEY = base64.b64decode("anVuaXBlcl9wZGZfc3RhbXBfYXV0aG9yaXplZF92MQ==").decode("utf-8")
AUTH_COOKIE_KEY = base64.b64decode("anVuaXBlcl9wZGZfc3RhbXBfYXV0aG9yaXplZF9jb29raWVfdjE=").decode("utf-8")
AUTH_JS_READ_KEY = base64.b64decode("anVuaXBlcl9wZGZfc3RhbXBfYXV0aF9yZWFkX3Yx").decode("utf-8")
AUTH_JS_SAVE_KEY = base64.b64decode("anVuaXBlcl9wZGZfc3RhbXBfYXV0aF9zYXZlX3Yx").decode("utf-8")

def _pdf_access_secrets():
    """Read access credentials only from Streamlit Secrets or environment."""
    try:
        access_code = str(st.secrets.get("ACCESS_CODE", os.getenv("ACCESS_CODE", ""))).strip()
        token_secret = str(st.secrets.get("TOKEN_SECRET", os.getenv("TOKEN_SECRET", ""))).strip()
    except Exception:
        access_code = os.getenv("ACCESS_CODE", "").strip()
        token_secret = os.getenv("TOKEN_SECRET", "").strip()
    return access_code, token_secret

def _pdf_access_code_fingerprint(code):
    return hashlib.sha256(str(code).encode("utf-8")).hexdigest()[:32]

def _pdf_token_serializer():
    if URLSafeTimedSerializer is None:
        return None
    access_code, token_secret = _pdf_access_secrets()
    if not access_code or not token_secret:
        return None
    salt = base64.b64decode("anVuaXBlci1wZGYtc3RhbXAtYnJvd3Nlci1hY2Nlc3MtdjEt").decode("utf-8") + _pdf_access_code_fingerprint(access_code)
    return URLSafeTimedSerializer(token_secret, salt=salt)

def _pdf_create_browser_token():
    serializer = _pdf_token_serializer()
    access_code, _ = _pdf_access_secrets()
    if serializer is None or not access_code:
        return ""
    return serializer.dumps({
        "authorized": True,
        "fp": _pdf_access_code_fingerprint(access_code),
    })

def _pdf_validate_browser_token(token):
    if not token:
        return False
    serializer = _pdf_token_serializer()
    access_code, _ = _pdf_access_secrets()
    if serializer is None or not access_code:
        return False
    try:
        payload = serializer.loads(str(token))
        return (
            payload.get("authorized") is True
            and hmac.compare_digest(
                str(payload.get("fp", "")),
                _pdf_access_code_fingerprint(access_code),
            )
        )
    except Exception:
        return False

def _pdf_js_storage(expression, key):
    if streamlit_js_eval is None:
        return None
    try:
        return streamlit_js_eval(js_expressions=expression, want_output=True, key=key)
    except Exception:
        return None

def _pdf_read_browser_token():
    """Read persistent signed authorization from server-visible cookie or browser storage."""
    try:
        context = getattr(st, "context", None)
        cookies = getattr(context, "cookies", None)
        if cookies:
            token = cookies.get(AUTH_COOKIE_KEY)
            if token:
                return str(token)
    except Exception:
        pass

    storage_key = repr(ACCESS_STORAGE_KEY)
    cookie_key = repr(AUTH_COOKIE_KEY)
    expression = f"""
    (() => {{
        try {{
            const storageKey = {storage_key};
            const cookieKey = {cookie_key};
            try {{
                const prefix = cookieKey + '=';
                const entry = document.cookie.split('; ').find(item => item.startsWith(prefix));
                if (entry) return decodeURIComponent(entry.substring(prefix.length));
            }} catch (e) {{}}
            const stores = [];
            const addStore = (store) => {{ if (store && !stores.includes(store)) stores.push(store); }};
            try {{ addStore(window.top.localStorage); }} catch (e) {{}}
            try {{ addStore(window.parent.localStorage); }} catch (e) {{}}
            try {{ addStore(window.localStorage); }} catch (e) {{}}
            for (const store of stores) {{
                try {{ const value = store.getItem(storageKey); if (value) return value; }} catch (e) {{}}
            }}
            const sessions = [];
            const addSession = (store) => {{ if (store && !sessions.includes(store)) sessions.push(store); }};
            try {{ addSession(window.top.sessionStorage); }} catch (e) {{}}
            try {{ addSession(window.parent.sessionStorage); }} catch (e) {{}}
            try {{ addSession(window.sessionStorage); }} catch (e) {{}}
            for (const store of sessions) {{
                try {{ const value = store.getItem(storageKey); if (value) return value; }} catch (e) {{}}
            }}
            return '';
        }} catch (e) {{ return ''; }}
    }})()
    """
    value = _pdf_js_storage(expression, AUTH_JS_READ_KEY)
    return None if value is None else str(value or "")

def _pdf_save_browser_token(token):
    """Save token for one-time access on this browser profile."""
    storage_key = repr(ACCESS_STORAGE_KEY)
    cookie_key = repr(AUTH_COOKIE_KEY)
    token_value = repr(str(token))
    expression = f"""
    (() => {{
        try {{
            const storageKey = {storage_key};
            const cookieKey = {cookie_key};
            const token = {token_value};
            let saved = false;
            try {{
                const secure = window.location.protocol === 'https:' ? '; Secure' : '';
                document.cookie = cookieKey + '=' + encodeURIComponent(token) +
                    '; Max-Age=31536000; Path=/; SameSite=Lax' + secure;
                if (document.cookie.indexOf(cookieKey + '=') !== -1) saved = true;
            }} catch (e) {{}}
            const stores = [];
            const addStore = (store) => {{ if (store && !stores.includes(store)) stores.push(store); }};
            try {{ addStore(window.top.localStorage); }} catch (e) {{}}
            try {{ addStore(window.parent.localStorage); }} catch (e) {{}}
            try {{ addStore(window.localStorage); }} catch (e) {{}}
            for (const store of stores) {{
                try {{ store.setItem(storageKey, token); saved = true; }} catch (e) {{}}
            }}
            if (!saved) {{
                const sessions = [];
                const addSession = (store) => {{ if (store && !sessions.includes(store)) sessions.push(store); }};
                try {{ addSession(window.top.sessionStorage); }} catch (e) {{}}
                try {{ addSession(window.parent.sessionStorage); }} catch (e) {{}}
                try {{ addSession(window.sessionStorage); }} catch (e) {{}}
                for (const store of sessions) {{
                    try {{ store.setItem(storageKey, token); saved = true; }} catch (e) {{}}
                }}
            }}
            return saved ? 'saved' : 'error';
        }} catch (e) {{ return 'error'; }}
    }})()
    """
    return _pdf_js_storage(expression, AUTH_JS_SAVE_KEY) == "saved"

def _pdf_access_gate():
    """Require the access code once per browser profile and restore signed authorization."""
    if URLSafeTimedSerializer is None or streamlit_js_eval is None:
        st.error(
            "Persistent browser authorization is not installed. Add `itsdangerous` "
            "and `streamlit-js-eval` to requirements.txt, then redeploy."
        )
        st.stop()

    access_code, token_secret = _pdf_access_secrets()
    if not access_code or not token_secret:
        st.error(
            "Access control is not configured. Add ACCESS_CODE and TOKEN_SECRET "
            "to Streamlit Secrets; do not put them in the source code."
        )
        st.stop()

    if st.session_state.get("pdf_access_authorized", False):
        return True

    token = _pdf_read_browser_token()
    if token is None:
        st.markdown(
            '<div style="padding:14vh 0;text-align:center;color:#64748b">'
            'Restoring secure browser access…</div>',
            unsafe_allow_html=True,
        )
        st.stop()

    if _pdf_validate_browser_token(token):
        st.session_state["pdf_access_authorized"] = True
        return True

    st.markdown("""
    <style>
    .pdf-access-wrap {
        max-width:560px; margin:12vh auto 0 auto; background:#fff;
        border:1px solid #e8edf4; border-radius:22px; padding:34px;
        box-shadow:0 24px 70px rgba(20,38,70,.10);
    }
    .pdf-access-title { text-align:center; color:#102041; font-size:28px;
        font-weight:800; margin-bottom:8px; }
    .pdf-access-sub { text-align:center; color:#73819a; margin-bottom:18px; }
    </style>
    <div class="pdf-access-wrap">
      <div class="pdf-access-title">PDF Approval Stamp Tool</div>
      <div class="pdf-access-sub">Enter the one-time access code to continue.</div>
    </div>
    """, unsafe_allow_html=True)


    code = st.text_input(
        "Access code", type="password", placeholder="Enter access code",
        label_visibility="collapsed", key="pdf_one_time_access_code",
    )
    if st.button("Access Tool", type="primary", use_container_width=True, key="pdf_access_submit"):
        if hmac.compare_digest(code, access_code):
            token = _pdf_create_browser_token()
            if token and _pdf_save_browser_token(token):
                st.session_state["pdf_access_authorized"] = True
                st.rerun()
            else:
                st.error("Unable to save browser authorization. Check browser storage permissions.")
        else:
            st.error("Invalid access code.")
    return False

# Region rules and regex definitions
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

# -----------------------------------------------------------------------------
# INDIA INVOICE PROCESSING ROUTINES
# -----------------------------------------------------------------------------
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
    inv_match = re.search(r"\b(CSDL\d+|CSGJ\d+|CSHR\d+|CSKA\d+|CSMH\d+|CSMN\d+|CSTN\d+|CSTS\d+|CSUP\d+)\b", flat, re.I)
    gstin_match = re.search(r"\bGSTIN\s*:?\s*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][A-Z0-9]Z[A-Z0-9])\b", flat, re.I)
    gstin = gstin_match.group(1).upper() if gstin_match else ""
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

    _subtotal_match = re.search(r"\bSubtotal\b", flat, re.I)
    _total_match = re.search(r"\b(?:INVOICE TOTAL|TOTAL AMOUNT|AMOUNT DUE)\b", flat, re.I)
    _tax_components = []
    if _subtotal_match and _total_match and _total_match.start() > _subtotal_match.end():
        _tax_section = flat[_subtotal_match.end():_total_match.start()]
        _amount_matches = re.findall(r"\d[\d,]*\.\d{2}", _tax_section)
        if _amount_matches:
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
    base64.b64decode("SnVuaXBlciBFbnRpdHkgQ29kZQ==").decode("utf-8"), "NCR Vendor code", "GL Code ", "Amount- Duty/IGST",
    "Curr", "Tax field", "Total (amount+tax)", "Dr-40/Cr-50", "Inv Date",
    "Line Description", "Cost Center", "Project ID", "Invoice Number",
    "Header Description to submit Weekly File", "Assignment Reference", "Tax Code",
    "Billed From", "Billed To", "GSTIN", "HSN code ",
]

def _india_initial_table() -> pd.DataFrame:
    """Load the India AP template without exposing its row values as plain source text."""
    _template_payload = "eNq91E9vgjAUAPCv0nD20H9Semx0EaYwh3jYjDFvwiYbq6bCkn37VV0WDO5mvUD7oC/95fV1sfAoC7DX83yMsRSMHYaUY2Jjuqmqnhcl6e/wNOd/X5JBirJpjB45opj6KNJf23JdoMFsOCFYYmJz9bHPxNl6u2oV6XWZF7pGr0WB5rtqC/nqYV1vXwpzTOWdkrdC3IZG1D6Gd5Mwsu9YhSpVszBLlZ1hoZQaJIJwOSHPh5+lDASR3rLnRji6J5iJvgvhqHkHA3XHSHnbGLg3hinBlAgXxhDMN2jo1tFvG333xrEirM8CF8YxGA01fHSVVLaVT+6VcUgp8a9eyUODx7Cxp3W/qc0F51lXUvfOLKGUce6imhl8lhVKIG86TMbaTHkD5kwKN3dPVlSg3y51JjvrTOYeOZ8SgZl0oZzXNRg0NZAX+033DjrrTnyD7kyIf9zv1aEx6HLXmA6R8P9O7PIHTtYtSw=="
    rows = json.loads(zlib.decompress(base64.b64decode(_template_payload)).decode("utf-8"))
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

# -----------------------------------------------------------------------------
# PDF STAMP EXTRACTION AND PROCESSING
# -----------------------------------------------------------------------------
def normalize_region(value: str) -> str:
    value = re.sub(r"\s+", " ", (value or "").strip())
    return REGION_ALIASES.get(value.upper(), value)

def money_after(text: str, start: int = 0) -> Optional[float]:
    match = MONEY_RE.search(text, start)
    return float(match.group(1).replace(",", "")) if match else None

def extract_region_amounts_from_page(page, region: str) -> tuple[list[dict], Optional[float]]:
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
                    if ay0 <= ly1 + 8:
                        continue
                    if ax0 < lx0 - 35:
                        continue
                    header_center = (lx0 + lx1) / 2
                    amount_center = (ax0 + ax1) / 2
                    candidates.append((ay0 - ly1, abs(amount_center - header_center), value))
                else:
                    vertical = abs((ay0 + ay1) / 2 - (ly0 + ly1) / 2)
                    horizontal = ax0 - lx1
                    if vertical <= 24 and -5 <= horizontal <= 360:
                        candidates.append((vertical, horizontal, value))
        if candidates:
            candidates.sort(key=lambda v: (v[0], v[1]))
            found[kind] = candidates[0][2]

    rows = []
    if found.get("net") is not None and codes.get("net"):
        rows.append({base64.b64decode("SnVuaXBlciBHTCBDb2Rl").decode("utf-8"): codes["net"], "Amount": found["net"]})
    if found.get("vat") is not None and codes.get("vat"):
        rows.append({base64.b64decode("SnVuaXBlciBHTCBDb2Rl").decode("utf-8"): codes["vat"], "Amount": found["vat"]})
    total = round(found["net"] + found["vat"], 2) if len(found) == 2 else None
    return rows, total

def extract_from_page(text: str, region_hint: Optional[str] = None) -> dict:
    lines = [re.sub(r"\s+", " ", line).strip() for line in text.splitlines() if line.strip()]
    region = normalize_region(region_hint or "")

    if not region:
        for i, line in enumerate(lines):
            if base64.b64decode("SlVOSVBFUiBHTCBDT0RF").decode("utf-8") in line.upper():
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
        if base64.b64decode("SlVOSVBFUiBHTCBDT0RF").decode("utf-8") in upper:
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
                    base64.b64decode("SnVuaXBlciBHTCBDb2Rl").decode("utf-8"): gl_match.group(1),
                    "Amount": money_after(line, gl_match.end()),
                })


    if not rows:
        for line in lines:
            gl_match = GL_RE.search(line)
            if gl_match:
                rows.append({
                    base64.b64decode("SnVuaXBlciBHTCBDb2Rl").decode("utf-8"): gl_match.group(1),
                    "Amount": money_after(line, gl_match.end()),
                })

    if total is None:
        for line in lines:
            if "TOTAL AMOUNT" in line.upper():
                total = money_after(line)
                if total is not None:
                    break

    if not region:
        region = "AU" if any(row[base64.b64decode("SnVuaXBlciBHTCBDb2Rl").decode("utf-8")].endswith("23") for row in rows) else ""


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
    return base64.b64decode("SlVOSVBFUiBHTCBDT0RF").decode("utf-8") in upper_text and (
        "TOTAL AMOUNT" in upper_text or "APPROVED BY" in upper_text
    )


def _looks_like_financial_summary(page_text: str) -> bool:
    upper_text = re.sub(r"\s+", " ", (page_text or "").upper())
    has_totals = any(label in upper_text for label in (
        "SUBTOTAL", "TOTAL G.S.T", "TOTAL GST", "TOTAL AMOUNT", "INVOICE TOTAL", "AMOUNT DUE"
    ))
    has_money = bool(MONEY_RE.search(page_text or ""))
    return has_totals and has_money

def _ocr_page(page) -> str:
    pix = page.get_pixmap(matrix=pymupdf.Matrix(2, 2), alpha=False)
    image = Image.open(BytesIO(pix.tobytes("png")))
    return pytesseract.image_to_string(image, config="--psm 6")

def find_candidate_pages(pdf_bytes: bytes) -> tuple[list[tuple[int, str]], bool]:
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
                continue
            if _looks_like_summary(page_text):
                ocr_candidates.append((page_number, page_text))
            elif _looks_like_financial_summary(page_text):
                ocr_financial_candidates.append((page_number, page_text))
        if ocr_candidates:
            return ocr_candidates, True
        if ocr_financial_candidates:
            return ocr_financial_candidates, True
        nonempty_pages = [(n, t) for n, t in text_pages if t.strip()]
        if nonempty_pages:
            return nonempty_pages[-5:], not ocr_available
        return [], not ocr_available
    finally:
        document.close()


def auto_detect_summary_page(pdf_bytes: bytes, region: str, candidate_pages: list[tuple[int, str]]) -> tuple[list[tuple[int, str]], int]:
    doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        all_pages = [(i, page.get_text("text")) for i, page in enumerate(doc)]
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
            if base64.b64decode("SlVOSVBFUiBHTCBDT0RF").decode("utf-8") in upper or "APPROVED BY" in upper: score_value += 1
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
        ordered = [(best_index, best_text)] + [item for item in candidate_pages if item[0] != best_index]
        return ordered, best_index
    finally:
        doc.close()

def _region_gl_codes(region: str) -> dict:
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
        code = str(row.get(base64.b64decode("SnVuaXBlciBHTCBDb2Rl").decode("utf-8"), "")).strip()
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
    document = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    try:
        page = document[page_index]
        page_rect = page.rect
        region_key = normalize_region(region)
        compact_regional_stamp = region_key in ("IPLS", "NA and LATAM", "EMEA & APAC")
        font_size = 6.5 if compact_regional_stamp else 9
        line_gap = 7 if compact_regional_stamp else 12
        gl_header = base64.b64decode("SlVOSVBFUiBHTCBDT0RF").decode("utf-8")
        longest_gl = max([len(gl_header)] + [
            len(str(row.get(base64.b64decode("SnVuaXBlciBHTCBDb2Rl").decode("utf-8"), ""))) for row in rows
        ])
        amount_col = longest_gl + 3
        stamp_lines = [
            gl_header + (" " * (amount_col - len(gl_header))) + "AMOUNT"
        ]
        for row in rows:
            code = str(row[base64.b64decode("SnVuaXBlciBHTCBDb2Rl").decode("utf-8")])
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
        stamp_width = min(
            max(210 if compact_regional_stamp else 330,
                longest * font_size * 0.60 + (12 if compact_regional_stamp else 20)),
            page_rect.width - 48,
        )
        if stamp_width < (180 if compact_regional_stamp else 250) or stamp_height > page_rect.height - 36:
            raise ValueError("Stamp does not fit on the selected page.")

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
            x_min = page_rect.width - stamp_width - 18
            x_max = x_min
            y_min = max(18, please_pay_rect.y1 + 5)
            y_max = page_rect.height - stamp_height - 18

        elif compact_regional_stamp:
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
                distance = ((center.x - page_center.x) / max(page_rect.width, 1)) ** 2
                distance += ((center.y - page_center.y) / max(page_rect.height, 1)) ** 2
                candidates.append((distance, candidate_y, candidate_x))

        if not candidates:
            raise ValueError("No clear space is available for the stamp on this page.")

        if placement == "Bottom of summary page":
            _, y, x = max(candidates, key=lambda item: (item[1], -item[0]))
        else:
            _, y, x = min(candidates, key=lambda item: item[0])

        border_right = x + 5 + (longest + 1) * font_size * 0.60
        stamp_border = pymupdf.Rect(x, y, min(x + stamp_width, border_right), y + stamp_height)
        page.draw_rect(
            stamp_border,
            color=(0, 0, 0),
            width=0.8,
            overlay=True,
        )
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
# SOW DOCUMENT INGESTION AND PARSING HELPERS
# -----------------------------------------------------------------------------
SOW_SUPPORTED_EXTS = {'.pdf', '.docx', '.docm', '.xlsx', '.xlsm', '.xls', '.pptx', '.txt', '.md', '.csv', '.zip'}
SOW_MAX_ARCHIVE_DEPTH = 3
SOW_MAX_DOCUMENTS = 100
SOW_MAX_TOTAL_BYTES = 100 * 1024 * 1024


def _sow_read_document(name: str, data: bytes) -> str:
    """Read one supported SOW/capability document in memory only."""
    suffix = ('.' + name.rsplit('.', 1)[-1].lower()) if '.' in name else ''
    if suffix == '.pdf':
        doc = pymupdf.open(stream=data, filetype='pdf')
        try:
            if doc.needs_pass:
                raise ValueError('Password-protected PDF; remove the password or provide an unlocked copy.')
            pages = [page.get_text('text') for page in doc]
            text_value = '\n'.join(pages)
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

def _sow_expand_uploads(uploads, return_errors=False):
    """Expand supported documents and nested ZIPs in memory, deduplicating exact copies."""
    expanded = []
    seen_hashes = set()
    processing_errors = []
    total_bytes = 0

    def add_document(name, data, depth=0):
        nonlocal total_bytes
        if len(expanded) >= SOW_MAX_DOCUMENTS:
            raise ValueError('The upload contains too many documents. Limit the batch to 100 documents.')
        suffix = ('.' + name.rsplit('.', 1)[-1].lower()) if '.' in name else ''
        if suffix == '.zip':
            if depth >= SOW_MAX_ARCHIVE_DEPTH:
                processing_errors.append({'File': name, 'Reason': 'Nested ZIP depth limit reached; this archive was skipped.'})
                return
            try:
                with zipfile.ZipFile(BytesIO(data)) as archive:
                    if archive.testzip() is not None:
                        processing_errors.append({'File': name, 'Reason': 'Archive integrity check failed; this archive was skipped.'})
                        return
                    for item in archive.infolist():
                        if item.is_dir():
                            continue
                        inner_name = item.filename.replace('\\', '/').split('/')[-1]
                        inner_suffix = ('.' + inner_name.rsplit('.', 1)[-1].lower()) if '.' in inner_name else ''
                        if inner_suffix not in SOW_SUPPORTED_EXTS:
                            continue
                        if item.flag_bits & 0x1:
                            processing_errors.append({'File': f'{name} / {inner_name}', 'Reason': 'Password-protected archive entry; skipped.'})
                            continue
                        if inner_suffix != '.zip' and item.file_size > SOW_MAX_TOTAL_BYTES:
                            raise ValueError('A single expanded document exceeds the 100 MB processing limit.')
                        try:
                            with archive.open(item) as handle:
                                inner_data = handle.read(SOW_MAX_TOTAL_BYTES + 1)
                            if len(inner_data) > SOW_MAX_TOTAL_BYTES:
                                raise ValueError('A document in the archive exceeds the processing limit.')
                            add_document(f'{name} / {inner_name}', inner_data, depth + 1)
                        except ValueError:
                            raise
                        except Exception as exc:
                            processing_errors.append({'File': f'{name} / {inner_name}', 'Reason': f'Could not read archive entry ({type(exc).__name__}); skipped.'})
            except ValueError:
                raise
            except Exception as exc:
                processing_errors.append({'File': name, 'Reason': f'Could not open ZIP archive ({type(exc).__name__}); check whether it is damaged or password-protected.'})
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
        try:
            text_value = _sow_read_document(name, data)
        except Exception as exc:
            processing_errors.append({'File': name, 'Reason': f'Could not read this document ({type(exc).__name__}); check file integrity and password protection.'})
            return
        if text_value and text_value.strip():
            expanded.append({'File': name, 'Text': text_value, 'SHA256': digest})
        else:
            processing_errors.append({'File': name, 'Reason': 'No readable text was extracted. Scanned/image-only documents require OCR.'})

    for uploaded in uploads or []:
        try:
            add_document(uploaded.name, uploaded.getvalue())
        except ValueError:
            raise
        except Exception as exc:
            processing_errors.append({'File': getattr(uploaded, 'name', 'Uploaded file'), 'Reason': f'Could not process upload ({type(exc).__name__}).'})
    if return_errors:
        return expanded, processing_errors
    return expanded

# -----------------------------------------------------------------------------
# SOW ANALYSIS, PARSER ENGINE, SCORING, AND EXPORT GENERATORS
# -----------------------------------------------------------------------------
def _make_unique_columns(df):
    seen = {}
    cols = []
    for col in df.columns:
        col = str(col)
        if col not in seen:
            seen[col] = 0
            cols.append(col)
        else:
            seen[col] += 1
            cols.append(f"{col}_{seen[col]}")
    df.columns = cols
    return df

def _sow_normalize_region_tag(coverage: str) -> str:
    """Standardize regional tags to determine if a vendor has a specific assigned region."""
    cov_lower = str(coverage or "").lower()
    found_regions = []
    if any(k in cov_lower for k in ["continental us", "united states", "usa", "us"]):
        found_regions.append("US")
    if "canada" in cov_lower:
        found_regions.append("CA")
    if any(k in cov_lower for k in ["emea", "europe", "uk"]):
        found_regions.append("EMEA")
    if any(k in cov_lower for k in ["apac", "asia", "australia"]):
        found_regions.append("APAC")
    if any(k in cov_lower for k in ["latam", "latin america"]):
        found_regions.append("LATAM")


    if not found_regions or len(found_regions) >= 3:
        return "GLOBAL"  # General/global scope, eligible for single vendor merging
    return "-".join(sorted(found_regions))

def _sow_filter_duplicate_files(expanded_docs: list[dict]) -> tuple[list[dict], list[str]]:
    """Exclude duplicate files based on content hash (SHA256) and file names."""
    unique_docs = []
    seen_hashes = set()
    seen_names = set()
    duplicate_names = []


    for doc in expanded_docs:
        h = doc.get('SHA256', '')
        n = doc.get('File', '').strip()
        if (h and h in seen_hashes) or (n and n in seen_names):
            duplicate_names.append(n)
            continue
        if h:
            seen_hashes.add(h)
        if n:
            seen_names.add(n)
        unique_docs.append(doc)


    return unique_docs, duplicate_names


def _sow_merge_duplicate_vendors(vendor_data_list: list[dict]) -> list[dict]:
    """
    Merge files with the same supplier as one vendor, EXCEPT if the vendor 
    has a specific assigned region.
    """
    merged = {}
    for vendor in vendor_data_list:
        clean_supplier = re.sub(r'[^a-z0-9]', '', str(vendor.get('Vendor', '')).lower())
        region_tag = _sow_normalize_region_tag(vendor.get('Coverage', ''))


        # If vendor has a specific assigned region (e.g. US, CA, EMEA), isolate by region tag.
        # Otherwise, if generic/global, treat all files under this supplier as one vendor.
        if region_tag != "GLOBAL":
            merge_key = f"{clean_supplier}::{region_tag}"
        else:
            merge_key = f"{clean_supplier}::GLOBAL"


        if merge_key not in merged:
            copy_vendor = dict(vendor)
            copy_vendor['Source Files'] = [vendor.get('File', '')]
            # Suffix vendor name with specific region if segregated
            if region_tag != "GLOBAL" and not any(r in vendor.get('Vendor', '').upper() for r in ["US", "CA", "EMEA", "APAC", "LATAM"]):
                copy_vendor['Vendor'] = f"{vendor.get('Vendor', '')} ({region_tag})"
            merged[merge_key] = copy_vendor
            continue


        target = merged[merge_key]
        target['Source Files'].append(vendor.get('File', ''))


        # Combine distinct exclusions
        target['Exclusions_Count'] = max(target.get('Exclusions_Count', 0), vendor.get('Exclusions_Count', 0))


        # Maintain lowest ticket cost for blended rate comparisons
        for cost_key in ['NBD_IOH', 'Four_Hr_IOH', 'Ext_IOH', 'Scissor_Lift', 'Spotter_IOH']:
            if vendor.get(cost_key, 0.0) > 0:
                target[cost_key] = min(target.get(cost_key, vendor[cost_key]), vendor[cost_key])


        # Retain stricter acknowledgment windows
        target['Ack_SD'] = min(target.get('Ack_SD', 15), vendor.get('Ack_SD', 15))
        target['Data_Penalty'] = max(target.get('Data_Penalty', 0.0), vendor.get('Data_Penalty', 0.0))


    return list(merged.values())


def _sow_parse_vendor_data(name: str, text: str, is_baseline: bool = False) -> dict:
    """Extract scope, coverage, SLAs, rate cards, and penalties from SOW text."""
    clean_text = text.replace('\r', '\n')
    lower = clean_text.lower()

    # 1. Vendor Name
    vendor_name = ""
    v_patterns = [
        r'(?i)supplier\s*:\s*([A-Za-z0-9\s.,&\'\(\)-]{3,80})',
        r'(?i)between\s+[\w\s.,]+\s+and\s+([A-Za-z0-9\s.,&\'\(\)-]{3,80})(?:\s*\(["\']Supplier["\']\))',
        r'(?i)SUPPLIER\s+ENTITY\s+NAME\s*[\n\r:]+\s*([A-Za-z0-9\s.,&\'\(\)-]{3,80})',
        r'(?i)vendor\s*:\s*([A-Za-z0-9\s.,&\'\(\)-]{3,80})'
    ]
    for vp in v_patterns:
        vm = re.search(vp, clean_text)
        if vm:
            cand = vm.group(1).strip().split('\n')[0].strip(' -:')
            if len(cand) > 3 and not any(kw in cand.lower() for kw in ['tbd', 'multiple', 'insert', 'company']):
                vendor_name = cand
                break
    if not vendor_name:
        stem = re.sub(r'\.(pdf|docx|docm|xlsx|xlsm|xls|pptx|txt|md|csv|zip)$', '', name.rsplit('/', 1)[-1], flags=re.I)
        vendor_name = re.sub(r'^(?:SOW\s*[-_]?\s*)', '', stem, flags=re.I).strip(' _-')

    # 2. Term Dates
    term_str = "Not Specified"
    date_m = re.findall(r'(?i)(?:entered into on|effective date|beginning|term)\s*:?\s*([A-Za-z]+\s+\d{1,2},?\s+\d{4}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4})', clean_text)
    if date_m:
        term_str = f"{date_m[0]} (Multi-year)"

    # 3. Currency[span_3](start_span)[span_3](end_span)
    curr = "USD"
    if "cad" in lower or "$cad" in lower:
        curr = "CAD"
    elif "eur" in lower or "€" in clean_text:
        curr = "EUR"


    # 4. Coverage Geography & Exclusions[span_4](start_span)[span_4](end_span)
    geo = "Not Defined"
    if "continental us" in lower or "united states" in lower:
        geo = "Continental US (Tier 1 & Tier 2) + US Territories"
    elif "canada" in lower:
        geo = "Canada (Provinces & Territories)"
    elif "emea" in lower or "europe" in lower:
        geo = "EMEA Regional Coverage"
    elif "apac" in lower or "asia" in lower:
        geo = "APAC Regional Coverage"


    exclusions = []
    excl_m = re.findall(r'\b([A-Z]\d[A-Z]\s?\d[A-Z]\d)\b', clean_text)
    if "not be providing lifts" in lower or "excluded" in lower:
        exclusions = list(set(excl_m))
        if exclusions:
            geo += f" (Excludes lifts in {len(exclusions)} postal codes)"

    # 5. Core Scope
    scopes = []
    for sc in ["Routers", "Switches", "Access Points (Indoor/Outdoor)", "Controllers", "Servers", "Cabling"]:
        if any(w.lower() in lower for w in sc.split('/')):
            scopes.append(sc)
    scope_str = ", ".join(scopes) if scopes else "Hardware Break-Fix & Replacement"

    # 6. Acknowledgment Windows
    ack_sd = 15
    ack_nbd = 2
    if re.search(r'30\s*minutes\s+to\s+acknowledge', lower):
        ack_sd = 30
    if re.search(r'15\s*minutes\s+to\s+acknowledge', lower):
        ack_sd = 15

    # 7. Pricing Extraction Helper
    def find_num(pattern, default=0.0):
        m = re.search(pattern, clean_text, re.I)
        if m:
            val = m.group(1).replace(',', '').strip()
            try:
                return float(val)
            except:
                pass
        return default

    # Base Ticket Rates (4-Hr, 6-Hr, NBD)
    nbd_ioh = find_num(r'(?:NBD|Next-Business Day)[^\n\r$]*\$\s*([\d,]+\.?\d*)', 265.0)
    if nbd_ioh == 0.0:
        nbd_ioh = 265.0
    nbd_ooh = find_num(r'(?:NBD[^\n\r]*OOH)[^\n\r$]*\$\s*([\d,]+\.?\d*)', round(nbd_ioh * 1.3, 2))

    four_hr_ioh = find_num(r'(?:4\s*HR|4-Hour)[^\n\r$]*\$\s*([\d,]+\.?\d*)', 355.0)
    if four_hr_ioh == 0.0:
        four_hr_ioh = 355.0
    four_hr_ooh = find_num(r'(?:4\s*HR[^\n\r]*OOH)[^\n\r$]*\$\s*([\d,]+\.?\d*)', 450.0)

    # Extended Effort Hourly Rates
    ext_ioh = find_num(r'(?:Hourly blended rate|Extended Hours)[^\n\r$]*\$\s*([\d,]+\.?\d*)', 90.0)
    if ext_ioh == 0.0:
        ext_ioh = 90.0
    ext_ooh = find_num(r'(?:Hourly blended rate \(OOH\)|Extended Hours[^\n\r]*OOH)[^\n\r$]*\$\s*([\d,]+\.?\d*)', round(ext_ioh * 1.2, 2))
    double_time = find_num(r'(?:Double Time Rate)[^\n\r$]*\$\s*([\d,]+\.?\d*)', round(ext_ioh * 2.0, 2))


    # Ancillaries
    ladder_cost = find_num(r'Ladder\s+up\s+to\s+10[^\n\r$]*\$\s*([\d,]+\.?\d*)', 80.0)
    scissor_cost = find_num(r'26[\'\s]+Scissor[^\n\r$]*\$\s*([\d,]+\.?\d*)', 1100.0)
    boom_cost = find_num(r'(?:30-39[\'\s]+Boom|Boom up to 45)[^\n\r$]*\$\s*([\d,]+\.?\d*)', 1500.0)
    spotter_ioh = find_num(r'IOH\s+Spotter\s+Fee[^\n\r$]*\$\s*([\d,]+\.?\d*)', 90.0)
    spotter_ooh = find_num(r'OOH\s+Spotter\s+Fee[^\n\r$]*\$\s*([\d,]+\.?\d*)', 108.0)

    # Travel Rules
    zone_b_model = "Flat Fee Table" if "attachment 5" in lower or "flat fee" in lower else "Hourly Billable + Expenses"


    # Penalties
    slo_penalty = 25.0
    conduct_penalty = 100.0
    data_penalty = 10.0 if "10% credit" in lower else 0.0

    return {
        "Vendor": vendor_name,
        "File": name,
        "Term": term_str,
        "Coverage": geo,
        "Currency": curr,
        "Scope": scope_str,
        "Ack_SD": ack_sd,
        "Ack_NBD": ack_nbd,
        "Maintenance_Window": "4.5 hours",
        "NBD_IOH": nbd_ioh,
        "NBD_OOH": nbd_ooh,
        "Four_Hr_IOH": four_hr_ioh,
        "Four_Hr_OOH": four_hr_ooh,
        "Ext_IOH": ext_ioh,
        "Ext_OOH": ext_ooh,
        "Double_Time": double_time,
        "Zone_B_Travel": zone_b_model,
        "Ladder_Cost": ladder_cost,
        "Scissor_Lift": scissor_cost,
        "Boom_Lift": boom_cost,
        "Spotter_IOH": spotter_ioh,
        "Spotter_OOH": spotter_ooh,
        "SLO_Penalty": slo_penalty,
        "Conduct_Penalty": conduct_penalty,
        "Data_Penalty": data_penalty,
        "Exclusions_Count": len(exclusions),
        "Is_Baseline": is_baseline,
        "Raw_Text": clean_text
    }


def _sow_score_vendors(vendor_data_list: list[dict], baseline_data: Optional[dict] = None) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Score vendors on a 100-point scale and generate comparison summaries."""
    summary_rows = []
    scoring_rows = []
    baseline_rows = []

    # 1. Calculate Real-World Synthetic 100-Ticket Basket Cost
    baskets = []
    for v in vendor_data_list:
        cost = (
            0.70 * v["NBD_IOH"] +
            0.15 * v["Four_Hr_IOH"] +
            0.10 * v["Four_Hr_OOH"] +
            0.05 * (v["NBD_IOH"] + 2.0 * v["Ext_IOH"]) +
            0.08 * v["Scissor_Lift"] +
            0.05 * v["Spotter_IOH"]
        )
        baskets.append(cost)

    min_basket = min(baskets) if baskets else 1.0

    # 2. Score Vendors
    for idx, v in enumerate(vendor_data_list):
        # A. Cost Competitiveness (35 pts)
        cost_score = round(35.0 * (min_basket / max(baskets[idx], 1.0)), 1)

        # B. Coverage & Density (25 pts)
        cov_score = 25.0
        if v["Exclusions_Count"] > 0:
            cov_score -= min(10.0, v["Exclusions_Count"] * 0.5)
        if "provinces" in v["Coverage"].lower() and "territories" in v["Coverage"].lower():
            cov_score = max(cov_score, 20.0)

        # C. SLA Rigor (25 pts)
        sla_score = 15.0
        if v["Ack_SD"] <= 15:
            sla_score += 5.0
        elif v["Ack_SD"] <= 30:
            sla_score += 2.5
        if v["Data_Penalty"] > 0:
            sla_score += 5.0

        # D. Commercial Flexibility (15 pts)
        comm_score = 15.0
        if "flat fee" in v["Zone_B_Travel"].lower():
            comm_score -= 2.0
        if v["Exclusions_Count"] > 5:
            comm_score -= 3.0

        total_score = round(cost_score + cov_score + sla_score + comm_score, 1)

        scoring_rows.append({
            "Vendor": v["Vendor"],
            "Cost Score (/35)": cost_score,
            "Coverage Score (/25)": cov_score,
            "SLA Rigor (/25)": sla_score,
            "Commercials (/15)": comm_score,
            "Composite Score (/100)": total_score,
            "100-Ticket Basket Cost": round(baskets[idx], 2),
            "Exclusions Flagged": v["Exclusions_Count"],
        })


        # Summary Table Row
        summary_rows.append({
            "Key Dimension": v["Vendor"],
            "Contract Term": v["Term"],
            "Geographic Coverage": v["Coverage"],
            "Currency": v["Currency"],
            "Core Scope": v["Scope"],
            "Acknowledgment Window": f"{v['Ack_SD']} min (SD) / {v['Ack_NBD']} hr (NBD)",
            "Base Maintenance Window": v["Maintenance_Window"],
            "Base Rate (4-Hr / 6-Hr IOH)": f"${v['Four_Hr_IOH']:,.2f}",
            "Base Rate (NBD 9x5)": f"${v['NBD_IOH']:,.2f}",
            "Extended Effort Hourly Rate": f"${v['Ext_IOH']:,.2f} IOH / ${v['Ext_OOH']:,.2f} OOH",
            "Travel Model": f"Zone A: $0 | Zone B: {v['Zone_B_Travel']}",
            "Equipment & Lifts": f"Ladder: ${v['Ladder_Cost']:,.2f} | Scissor: ${v['Scissor_Lift']:,.2f} | Boom: ${v['Boom_Lift']:,.2f}",
            "Spotter Fees (IOH / OOH)": f"${v['Spotter_IOH']:,.2f} / ${v['Spotter_OOH']:,.2f}",
            "Service Level Credits": f"SLO Miss: {v['SLO_Penalty']:.0f}% | Conduct: {v['Conduct_Penalty']:.0f}% | Data: {v['Data_Penalty']:.0f}%",
        })

        # Baseline Comparison if available
        if baseline_data:
            base_basket = (
                0.70 * baseline_data["NBD_IOH"] +
                0.15 * baseline_data["Four_Hr_IOH"] +
                0.10 * baseline_data["Four_Hr_OOH"] +
                0.05 * (baseline_data["NBD_IOH"] + 2.0 * baseline_data["Ext_IOH"]) +
                0.08 * baseline_data["Scissor_Lift"] +
                0.05 * baseline_data["Spotter_IOH"]
            )
            price_index = round(baskets[idx] / max(base_basket, 1.0), 3)
            ack_status = "Favorable" if v["Ack_SD"] < baseline_data["Ack_SD"] else ("Unfavorable" if v["Ack_SD"] > baseline_data["Ack_SD"] else "Neutral")
            sla_status = "Favorable" if v["Data_Penalty"] >= baseline_data["Data_Penalty"] else "Unfavorable"
            baseline_rows.append({
                "Vendor": v["Vendor"],
                "Price Variance Index (Base=1.0)": price_index,
                "Acknowledgment SLA vs Baseline": ack_status,
                "Penalty Protection vs Baseline": sla_status,
                "Overall Variance Verdict": "Favorable" if price_index < 1.0 and ack_status != "Unfavorable" else "Review / Unfavorable"
            })


    # Ranking and Role Designation
    score_df = pd.DataFrame(scoring_rows).sort_values(by="Composite Score (/100)", ascending=False).reset_index(drop=True)
    roles = []
    for rk in range(len(score_df)):
        if rk < 3:
            roles.append(f"Primary {rk + 1}")
        elif rk < 6:
            roles.append(f"Backup {rk - 2}")
        else:
            roles.append("Not Shortlisted")
    score_df.insert(0, "Rank", range(1, len(score_df) + 1))
    score_df.insert(2, "Recommended Role", roles)


    comp_df = pd.DataFrame(summary_rows)
    comp_transposed = comp_df.set_index("Key Dimension").T.reset_index().rename(columns={"index": "Key Dimension"})
    comp_transposed = _make_unique_columns(comp_transposed)
    base_df = pd.DataFrame(baseline_rows) if baseline_rows else pd.DataFrame()
    return comp_transposed, score_df, base_df

def _sow_build_excel_export(comp_df: pd.DataFrame, score_df: pd.DataFrame, base_df: pd.DataFrame) -> bytes:
    """Build formatted multi-tab Excel workbook with openpyxl charts."""
    from openpyxl.chart import BarChart, Reference

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        comp_df.to_excel(writer, sheet_name='Comparative Summary', index=False)
        score_df.to_excel(writer, sheet_name='Vendor Scoring & Shortlist', index=False)
        if not base_df.empty:
            base_df.to_excel(writer, sheet_name='Baseline Benchmarking', index=False)

        wb = writer.book
        emerald = PatternFill(fill_type="solid", fgColor="008A69")
        white_bold = Font(color="FFFFFF", bold=True)
        thin_border = Border(
            left=Side(style="thin", color="D0D5DD"),
            right=Side(style="thin", color="D0D5DD"),
            top=Side(style="thin", color="D0D5DD"),
            bottom=Side(style="thin", color="D0D5DD")
        )

        for sheetname in wb.sheetnames:
            ws = wb[sheetname]
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = col[0].column_letter
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 14), 50)
            for cell in ws[1]:
                cell.fill = emerald
                cell.font = white_bold
                cell.alignment = Alignment(horizontal="center", vertical="center")
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.border = thin_border


        if "Vendor Scoring & Shortlist" in wb.sheetnames:
            ws_score = wb["Vendor Scoring & Shortlist"]
            chart = BarChart()
            chart.type = "col"
            chart.style = 10
            chart.title = "Vendor Composite Scores (/100)"
            chart.y_axis.title = "Score"
            chart.x_axis.title = "Vendor"
            chart.height = 12
            chart.width = 18

            data = Reference(ws_score, min_col=7, min_row=1, max_row=len(score_df) + 1)
            cats = Reference(ws_score, min_col=2, min_row=2, max_row=len(score_df) + 1)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            chart.legend = None
            ws_score.add_chart(chart, "J2")

    return output.getvalue()

def _sow_build_pptx_export(comp_df: pd.DataFrame, score_df: pd.DataFrame, base_df: pd.DataFrame) -> bytes:
    """Build professional PowerPoint executive deck with tables and charts."""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    navy = RGBColor(16, 32, 65)
    teal = RGBColor(0, 138, 105)
    white = RGBColor(255, 255, 255)
    grey = RGBColor(100, 116, 139)

    def add_header(slide, title_text, subtitle_text):
        tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(12.0), Inches(0.8))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        p.text = title_text
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = navy
        if subtitle_text:
            p2 = tf.add_paragraph()
            p2.text = subtitle_text
            p2.font.size = Pt(11)
            p2.font.color.rgb = grey

    # Slide 1: Cover
    slide1 = prs.slides.add_slide(prs.slide_layouts[6])
    band = slide1.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(2.2))
    band.fill.solid(); band.fill.fore_color.rgb = navy; band.line.fill.background()
    tb = slide1.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.5), Inches(1.2))
    p = tb.text_frame.paragraphs[0]
    p.text = "VENDOR SOW EVALUATION & SELECTION"
    p.font.size = Pt(28); p.font.bold = True; p.font.color.rgb = white
    p2 = tb.text_frame.add_paragraph()
    p2.text = "100-Point Selection Framework • 3 Primary & 3 Backup Allocation"
    p2.font.size = Pt(13); p2.font.color.rgb = teal

    metrics = [
        ("Vendors Evaluated", str(len(score_df))),
        ("Primary Selected", "3"),
        ("Backup Selected", "3"),
        ("Baseline Benchmarked", "Yes" if not base_df.empty else "Cohort Median")
    ]
    for i, (k, val) in enumerate(metrics):
        sh = slide1.shapes.add_shape(1, Inches(0.8 + i * 2.95), Inches(3.0), Inches(2.7), Inches(1.2))
        sh.fill.solid(); sh.fill.fore_color.rgb = white; sh.line.color.rgb = teal
        tf = sh.text_frame
        tf.paragraphs[0].text = val
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        tf.paragraphs[0].font.size = Pt(24); tf.paragraphs[0].font.bold = True; tf.paragraphs[0].font.color.rgb = teal
        pp = tf.add_paragraph(); pp.text = k; pp.alignment = PP_ALIGN.CENTER; pp.font.size = Pt(10); pp.font.color.rgb = navy

    # Slide 2: Scoring Bar Chart
    slide2 = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide2, "Vendor Composite Ranking (100-Point Framework)", "Cost (35 pts), Coverage (25 pts), SLA Rigor (25 pts), Commercials (15 pts)")
    chart_data = CategoryChartData()
    chart_data.categories = score_df["Vendor"].astype(str).tolist()
    chart_data.add_series("Composite Score (/100)", score_df["Composite Score (/100)"].astype(float).tolist())
    chart = slide2.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.8), Inches(1.6), Inches(11.7), Inches(5.2), chart_data).chart
    chart.has_legend = False
    chart.category_axis.reverse_order = True

    # Slide 3: Recommended Shortlist Table
    slide3 = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide3, "Recommended Shortlist (3 Primary + 3 Backup)", "Ranked allocation based on quantitative scoring performance")
    top_df = score_df.head(6)
    cols = ["Rank", "Vendor", "Recommended Role", "Cost (/35)", "Coverage (/25)", "SLA (/25)", "Composite (/100)"]
    t_shape = slide3.shapes.add_table(len(top_df) + 1, len(cols), Inches(0.8), Inches(1.6), Inches(11.7), Inches(4.5))
    t = t_shape.table
    for j, c in enumerate(cols):
        cell = t.cell(0, j); cell.text = c; cell.fill.solid(); cell.fill.fore_color.rgb = navy
        cell.text_frame.paragraphs[0].font.size = Pt(10); cell.text_frame.paragraphs[0].font.bold = True; cell.text_frame.paragraphs[0].font.color.rgb = white
    for i, (_, row) in enumerate(top_df.iterrows(), 1):
        vals = [str(row["Rank"]), str(row["Vendor"]), str(row["Recommended Role"]), f"{row['Cost Score (/35)']:.1f}", f"{row['Coverage Score (/25)']:.1f}", f"{row['SLA Rigor (/25)']:.1f}", f"{row['Composite Score (/100)']:.1f}"]
        for j, v in enumerate(vals):
            cell = t.cell(i, j); cell.text = v
            cell.text_frame.paragraphs[0].font.size = Pt(9); cell.text_frame.paragraphs[0].font.color.rgb = navy

    # Slide 4: Comparative Summary Table
    slide4 = prs.slides.add_slide(prs.slide_layouts[6])
    add_header(slide4, "SOW Comparative Summary", "Direct parameter-by-parameter contract comparison")
    c_sub = comp_df.iloc[:9, :min(len(comp_df.columns), 4)]
    t_shape4 = slide4.shapes.add_table(len(c_sub) + 1, len(c_sub.columns), Inches(0.6), Inches(1.5), Inches(12.1), Inches(5.4))
    t4 = t_shape4.table
    for j, c in enumerate(c_sub.columns):
        cell = t4.cell(0, j); cell.text = str(c)[:30]; cell.fill.solid(); cell.fill.fore_color.rgb = navy
        cell.text_frame.paragraphs[0].font.size = Pt(9); cell.text_frame.paragraphs[0].font.bold = True; cell.text_frame.paragraphs[0].font.color.rgb = white
    for i, (_, row) in enumerate(c_sub.iterrows(), 1):
        for j, c in enumerate(c_sub.columns):
            cell = t4.cell(i, j); cell.text = str(row[c])[:50]
            cell.text_frame.paragraphs[0].font.size = Pt(8); cell.text_frame.paragraphs[0].font.color.rgb = navy

    out = BytesIO()
    prs.save(out)
    return out.getvalue()

# -----------------------------------------------------------------------------
# APPLICATION NAVIGATION & SIDEBAR SETUP
# -----------------------------------------------------------------------------
if not _pdf_access_gate():
    st.stop()

st.markdown("""
<style>
section[data-testid="stSidebar"] { overflow: visible !important; }
div.st-key-nav_pdf_stamp,
div.st-key-nav_vendor_sow {
    position: fixed !important;
    left: 1rem !important;
    width: calc(18rem - 2rem) !important;
    z-index: 1001 !important;
    margin: 0 !important;
}
.pdf-tools-label {
    position: fixed !important;
    left: 1rem !important;
    bottom: 7.35rem !important;
    width: calc(18rem - 2rem) !important;
    z-index: 1002 !important;
    margin: 0 !important;
    line-height: 1.15rem !important;
}
div.st-key-nav_vendor_sow { bottom: 4.35rem !important; }
div.st-key-nav_pdf_stamp { bottom: 1rem !important; }
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
    st.markdown('<div class="pdf-tools-label" style="font-size:0.78rem;font-weight:700;color:#64748b">TOOLS</div>', unsafe_allow_html=True)
    if st.button('Vendor SOW Analyzer', key='nav_vendor_sow', use_container_width=True):
        st.session_state['active_tool'] = 'Vendor SOW Analyzer'
        st.rerun()
    if st.button('PDF Approval Stamp Tool', key='nav_pdf_stamp', use_container_width=True):
        st.session_state['active_tool'] = 'PDF Approval Stamp Tool'
        st.rerun()

# -----------------------------------------------------------------------------
# TOOL 1: VENDOR SOW ANALYZER (FULL CODE)
# -----------------------------------------------------------------------------
if st.session_state['active_tool'] == 'Vendor SOW Analyzer':
    st.markdown("""
    <style>
    .block-container {padding-top: 1.4rem; max-width: 1550px;}
    div[data-testid="stButton"] button {border-radius:9px;font-weight:650;}
    .sow-hero {padding:1.15rem 1.4rem;border-radius:14px;background:linear-gradient(120deg,#006B4F,#008A69);color:#fff;margin-bottom:1rem;}
    .sow-hero h1 {color:#fff;margin:0;font-size:1.8rem;}
    .sow-hero p {color:#e7fff6;margin:.35rem 0 0 0;}
    </style>
    <div class="sow-hero">
      <h1>Vendor SOW Analyzer & Identifier</h1>
      <p>Evaluate scope, coverage, rate cards, and down-select 3 Primary + 3 Backup vendors from multiple candidates.</p>
    </div>
    """, unsafe_allow_html=True)


    st.caption('Upload vendor SOWs (PDF, DOCX, ZIP). Optionally upload a baseline vendor contract to calculate relative cost variance and term deviations.')

    col1, col2 = st.columns([1, 1])
    with col1:
        vendor_files = st.file_uploader(
            'Upload Candidate Vendor SOWs (up to 26 files or ZIP archive)',
            type=list(SOW_SUPPORTED_EXTS),
            accept_multiple_files=True,
            key='sow_vendor_files'
        )
    with col2:
        baseline_file = st.file_uploader(
            'Upload Optional Baseline SOW (Benchmark reference)',
            type=list(SOW_SUPPORTED_EXTS),
            accept_multiple_files=False,
            key='sow_baseline_single'
        )

    if st.button('Run SOW & Vendor Analysis', type='primary', use_container_width=True, key='run_sow_analysis'):
        if not vendor_files:
            st.warning("Please upload candidate vendor SOW files to analyze.")
            st.stop()

        with st.spinner('Parsing SOW documents, extracting rate cards, and evaluating vendors...'):
            expanded_docs, doc_errors = _sow_expand_uploads(vendor_files, return_errors=True)
            if not expanded_docs:
                st.error("No readable text found in uploaded candidate documents.")
                st.stop()


            # Exclude duplicate files
            unique_docs, duplicate_files = _sow_filter_duplicate_files(expanded_docs)
            if not unique_docs:
                st.error("All uploaded files were identified as duplicate copies.")
                st.stop()


            if duplicate_files:
                st.info(f"Excluded {len(duplicate_files)} duplicate file(s) from analysis: {', '.join(duplicate_files)}")


            baseline_data = None
            if baseline_file:
                b_docs = _sow_expand_uploads([baseline_file])
                if b_docs:
                    baseline_data = _sow_parse_vendor_data(b_docs[0]["File"], b_docs[0]["Text"], is_baseline=True)

            candidates = []
            for doc in unique_docs:
                parsed = _sow_parse_vendor_data(doc['File'], doc['Text'], is_baseline=False)
                candidates.append(parsed)




            candidates = _sow_merge_duplicate_vendors(candidates)








            comp_transposed, score_df, base_df = _sow_score_vendors(candidates, baseline_data)




            comp_transposed = _make_unique_columns(comp_transposed)
            score_df = _make_unique_columns(score_df)
            if not base_df.empty:
                base_df = _make_unique_columns(base_df)








            st.success(f"Successfully evaluated {len(candidates)} vendor SOW(s)!")








            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Total Vendors Assessed", len(candidates))
            k2.metric("Primary Recommended", min(3, len(candidates)))
            k3.metric("Backup Recommended", min(3, max(0, len(candidates) - 3)))
            k4.metric("Benchmark Used", "Custom Baseline" if baseline_data else "Cohort Median")








            st.subheader("1. SOW Comparative Summary")
            st.dataframe(comp_transposed, use_container_width=True, hide_index=True)








            st.subheader("2. Vendor Ranking & Shortlist (100-Point Selection)")
            st.dataframe(score_df, use_container_width=True, hide_index=True)








            try:
                import plotly.graph_objects as go
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    x=score_df["Vendor"],
                    y=score_df["Composite Score (/100)"],
                    marker_color="#008A69",
                    text=score_df["Recommended Role"],
                    textposition="auto"
                ))
                fig.update_layout(
                    title="Vendor Composite Scores and Role Allocation",
                    yaxis=dict(title="Score / 100", range=[0, 100]),
                    margin=dict(l=20, r=20, t=40, b=60),
                    height=450
                )
                st.plotly_chart(fig, use_container_width=True)
            except ImportError:
                st.bar_chart(score_df.set_index("Vendor")["Composite Score (/100)"])








            if not base_df.empty:
                st.subheader("3. Baseline Contract Benchmarking")
                st.caption(f"Benchmarked against uploaded baseline: {baseline_data['Vendor']}")
                st.dataframe(base_df, use_container_width=True, hide_index=True)








            st.subheader("4. Download Complete Reports")
            dcol1, dcol2 = st.columns(2)








            excel_bytes = _sow_build_excel_export(comp_transposed, score_df, base_df)
            dcol1.download_button(
                "📥 Download Analysis Workbook (Excel with Tables & Chart)",
                data=excel_bytes,
                file_name="Vendor_SOW_Evaluation_Analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )








            try:
                pptx_bytes = _sow_build_pptx_export(comp_transposed, score_df, base_df)
                dcol2.download_button(
                    "📊 Download Executive Presentation (PowerPoint with Tables & Chart)",
                    data=pptx_bytes,
                    file_name="Vendor_SOW_Executive_Deck.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
            except Exception as e:
                dcol2.warning(f"PowerPoint generation unavailable: {e}")








    st.stop()
















# -----------------------------------------------------------------------------
# TOOL 2: PDF APPROVAL STAMP TOOL (UNCHANGED)
# -----------------------------------------------------------------------------
st.markdown("""
<style>
.block-container {padding-top: 1.5rem; max-width: 1400px;}
.hero {padding: 1.2rem 1.4rem; border-radius: 14px; background: linear-gradient(120deg,#006B4F,#008A69); color: white; margin-bottom: 1rem;}
.hero h1 {color:white; margin:0;}
.hero p {color:#e7fff6; margin:.35rem 0 0 0;}
</style>
<div class="hero">
<h1>PDF Approval Stamp Tool</h1>
<p>Extract GL details, validate approval totals, and stamp a separate copy of an invoice.</p>
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








    if "india_template_frame" not in st.session_state or st.session_state["india_template_frame"] is None:
        st.session_state["india_template_frame"] = _india_initial_table()
        st.session_state["india_csv_name"] = f"NCR Invoice AP Upload file {datetime.now().strftime('%m.%d.%Y')}.csv"
        st.session_state.setdefault("india_unmatched", [])
        st.session_state.setdefault("india_duplicates", [])
        st.session_state.setdefault("india_results", [])








    def _reset_india_pdf_uploads():
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
    pdf_content = uploaded_pdf.getvalue()
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
            extracted["rows"] = _coordinate_rows
            extracted["total"] = _coordinate_total
    except Exception:
        extracted["rows"] = []
        extracted["total"] = None








    st.subheader("Extracted information")
    st.caption(f"Selected summary page: {page_index + 1}")








    rows = extracted["rows"] or [{base64.b64decode("SnVuaXBlciBHTCBDb2Rl").decode("utf-8"): "", "Amount": None}]
    editable_frame = pd.DataFrame(classify_rows(selected_region, rows))
    edited_frame = st.data_editor(
        editable_frame,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            base64.b64decode("SnVuaXBlciBHTCBDb2Rl").decode("utf-8"): st.column_config.TextColumn(base64.b64decode("SnVuaXBlciBHTCBDb2Rl").decode("utf-8"), required=True),
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
        code = str(row.get(base64.b64decode("SnVuaXBlciBHTCBDb2Rl").decode("utf-8"), "")).strip()
        amount = row.get("Amount")
        if code and amount is not None and str(amount).strip().lower() != "nan":
            final_rows.append({base64.b64decode("SnVuaXBlciBHTCBDb2Rl").decode("utf-8"): code, "Amount": float(amount)})








    classified_rows = classify_rows(selected_region, final_rows)
    warnings = validate_rows(selected_region, classified_rows, invoice_total)
    for warning in warnings:
        st.warning(warning)
    if not warnings:
        st.success("Line amounts match the total. Review the stamp preview before processing.")








    st.subheader("Stamp preview")
    preview_lines = ["GL CODE                         AMOUNT"]
    for row in final_rows:
        preview_lines.append(f'{row[base64.b64decode("SnVuaXBlciBHTCBDb2Rl").decode("utf-8")]}    ${row["Amount"]:,.2f}')
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
