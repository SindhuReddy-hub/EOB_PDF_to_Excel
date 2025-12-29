import pdfplumber
import pytesseract
import pandas as pd
import cv2
import numpy as np
import re

# ------------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------------

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

EOB_TYPES = [
    "GEICO",
    "Provider Payment",
    "Offer of Payment",
    "Check",
    "Auto"
]

EXCEL_COLUMNS = [
    "eob_type",
    "payer_name",
    "patient_name",
    "claim_number",
    "invoice_number",
    "check_number",
    "check_date",
    "service_from",
    "service_to",
    "billed_amount",
    "allowed_amount",
    "paid_amount",
    "deductible",
    "copay",
    "page_number",
    "raw_text"
]

# ------------------------------------------------------------------
# COMMON HELPERS
# ------------------------------------------------------------------

def find(text, pattern, group=1):
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(group).strip() if match else None


def extract_check_number(text):
    primary = re.search(
        r"check\s*(number|no)\s*[:\-]?\s*([\w\-]+)",
        text,
        re.IGNORECASE
    )
    if primary:
        return primary.group(2).strip()

    fallback = re.search(
        r"NO\.?\s*N\s*(2\d{8})",
        text,
        re.IGNORECASE
    )
    if fallback:
        return fallback.group(1).strip()

    return None

# ------------------------------------------------------------------
# PAGE EXTRACTION
# ------------------------------------------------------------------

def extract_page_text(pdf_path):
    pages_text = []

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()

            if not text or len(text.strip()) < 50:
                img = page.to_image(resolution=300).original
                img = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
                text = pytesseract.image_to_string(img)

            pages_text.append({
                "page": i + 1,
                "text": text
            })

    return pages_text

# ------------------------------------------------------------------
# EOB TYPE DETECTION
# ------------------------------------------------------------------

def detect_eob_type(text):
    t = text.lower()

    if "geico" in t:
        return "GEICO"
    elif "provider payment" in t:
        return "Provider Payment"
    elif "offer of payment" in t:
        return "Offer of Payment"
    elif "check number" in t or "pay to the order of" in t:
        return "Check"
    else:
        return "Unknown"

# ------------------------------------------------------------------
# TYPE-SPECIFIC EXTRACTORS
# ------------------------------------------------------------------

def extract_geico_fields(text):
    return {
        "payer_name": "GEICO",
        "claim_number": find(text, r"claim\s*#?\s*[:\-]?\s*([0-9]{6,})"),
        "check_number": extract_check_number(text),
        "check_date": find(text, r"date\s*[:\-]?\s*([\d/]+)"),
        "paid_amount": find(text, r"total\s*amount\s*[:\-]?\s*\$?\**([\d,.]+)")
    }


def extract_provider_payment_fields(text):
    return {
        "payer_name": find(text, r"(insurance|health|mutual)[^\n]*"),
        "claim_number": find(text, r"claim\s*#?\s*[:\-]?\s*([0-9]{6,})"),
        "paid_amount": find(text, r"amount\s*paid\s*[:\-]?\s*\$?([\d,.]+)")
    }


def extract_offer_payment_fields(text):
    return {
        "invoice_number": find(text, r"invoice\s*(number|no)\s*[:\-]?\s*([\w\-]+)", 2),
        "paid_amount": find(text, r"amount\s*offered\s*[:\-]?\s*\$?([\d,.]+)")
    }


def extract_check_fields(text):
    return {
        "check_number": extract_check_number(text),
        "check_date": find(text, r"date\s*[:\-]?\s*([\d/]+)"),
        "paid_amount": find(text, r"\$([\d,.]+)")
    }

# ------------------------------------------------------------------
# ROUTER (THIS IS THE KEY)
# ------------------------------------------------------------------

def extract_by_eob_type(text, eob_type):

    if eob_type == "GEICO":
        return extract_geico_fields(text)

    elif eob_type == "Provider Payment":
        return extract_provider_payment_fields(text)

    elif eob_type == "Offer of Payment":
        return extract_offer_payment_fields(text)

    elif eob_type == "Check":
        return extract_check_fields(text)

    else:
        detected = detect_eob_type(text)
        return extract_by_eob_type(text, detected)

# ------------------------------------------------------------------
# PROCESS PAGES
# ------------------------------------------------------------------

def process_pages(pages, selected_eob_type):

    records = []

    for page in pages:
        text = page["text"]
        detected_type = detect_eob_type(text)

        if selected_eob_type != "Auto" and detected_type != selected_eob_type:
            continue

        record = {
            "eob_type": detected_type,
            "page_number": page["page"],
            "raw_text": text
        }

        record.update(extract_by_eob_type(text, detected_type))
        records.append(record)

    return records

# ------------------------------------------------------------------
# EXPORT
# ------------------------------------------------------------------

def export_to_excel(records, output_file):
    df = pd.DataFrame(records)
    df = df.reindex(columns=EXCEL_COLUMNS)
    df.to_excel(output_file, index=False)
