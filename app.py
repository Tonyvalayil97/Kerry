#!/usr/bin/env python3
# Streamlit UI – KLN Freight Invoice Extractor (Final Version)

import io
import os
import re
import traceback
from datetime import datetime
from typing import Dict, Any, Optional, List

import pdfplumber
import pandas as pd
import streamlit as st
from openpyxl import Workbook

# --------------------------------------------------------------
# Extract invoice number like DN26693 or DN26693A
# --------------------------------------------------------------
def extract_invoice_id(filename: str):
    m = re.search(r"(DN\s*\d+[A-Z]?)", filename.upper())
    if m:
        return m.group(1).replace(" ", "")
    return filename


# --------------------------------------------------------------
# REQUIRED COLUMNS (Your 13-column output)
# --------------------------------------------------------------
HEADERS = [
    "Timestamp", "Filename", "Invoice_Date", "Currency", "Shipper",
    "Weight_KG", "Volume_M3", "Chargeable_KG", "Chargeable_CBM",
    "Pieces", "Subtotal", "Freight_Mode", "Freight_Rate"
]


# --------------------------------------------------------------
# REGEX for KLN Freight Invoice Format
# --------------------------------------------------------------
INVOICE_DATE_PAT = re.compile(
    r"INVOICE DATE[\s:\-A-Z\n]*?(\d{4}-\d{2}-\d{2})", re.I
)

SHIPPER_PAT = re.compile(
    r"SHIPPER'S NAME\s*-\s*NOM DE L'EXP[ÉE]DITEUR\s*([\w\s\-\.,/&]+)", re.I
)

PACKAGES_PAT = re.compile(r"(\d+)\s+PACKAGE\b", re.I)

WEIGHT_PAT = re.compile(r"Gross Weight[:\s]+([\d.]+)\s*KG", re.I)
VOL_PAT = re.compile(r"Volume Weight[:\s]+([\d.]+)\s*KG", re.I)

SUBTOTAL_PAT = re.compile(
    r"([\d.,]+)\s*USD\s*Total", re.I
)

CURRENCY_PAT = re.compile(r"\b(USD|CAD|EUR)\b", re.I)

FREIGHT_RATE_PAT = re.compile(
    r"AIR FREIGHT.*?USD\s*([\d.,]+)", re.I
)


# --------------------------------------------------------------
# PDF PARSER (final)
# --------------------------------------------------------------
def parse_invoice_pdf_bytes(data: bytes, filename: str) -> Optional[Dict[str, Any]]:
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            text = "\n".join(p.extract_text() or "" for p in pdf.pages)

        # -------- Invoice Date --------
        inv_date = None
        m = INVOICE_DATE_PAT.search(text)
        if m:
            inv_date = m.group(1).strip()

        # -------- Currency --------
        currency = None
        m = CURRENCY_PAT.search(text)
        if m:
            currency = m.group(1).upper()

        # -------- Shipper --------
        shipper = None
        m = SHIPPER_PAT.search(text)
        if m:
            shipper = m.group(1).strip()

        # -------- Pieces --------
        pieces = None
        m = PACKAGES_PAT.search(text)
        if m:
            pieces = int(m.group(1))

        # -------- Weight KG --------
        weight = None
        m = WEIGHT_PAT.search(text)
        if m:
            weight = float(m.group(1))

        # -------- Volume Weight (KG) → m³ --------
        volume_m3 = None
        m = VOL_PAT.search(text)
        if m:
            vol_kg = float(m.group(1))
            volume_m3 = vol_kg / 167.0  # industry conversion

        # -------- Chargeable KG --------
        chargeable_kg = None
        if weight and volume_m3:
            chargeable_kg = max(weight, volume_m3 * 167)

        # -------- Chargeable CBM --------
        chargeable_cbm = volume_m3

        # -------- Freight Rate --------
        f_mode = None
        f_rate = None
        m = FREIGHT_RATE_PAT.search(text)
        if m:
            f_mode = "Air"
            f_rate = float(m.group(1).replace(",", ""))

        # -------- Subtotal --------
        subtotal = None
        m = SUBTOTAL_PAT.search(text)
        if m:
            subtotal = float(m.group(1).replace(",", ""))

        # -------- Return row --------
        return {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Filename": filename,
            "Invoice_Date": inv_date,
            "Currency": currency,
            "Shipper": shipper,
            "Weight_KG": weight,
            "Volume_M3": volume_m3,
            "Chargeable_KG": chargeable_kg,
            "Chargeable_CBM": chargeable_cbm,
            "Pieces": pieces,
            "Subtotal": subtotal,
            "Freight_Mode": f_mode,
            "Freight_Rate": f_rate,
        }

    except Exception:
        traceback.print_exc()
        return None


# --------------------------------------------------------------
# STREAMLIT INTERFACE
# --------------------------------------------------------------
st.set_page_config(
    page_title="KLN Invoice Extractor",
    page_icon="📄",
    layout="wide",
)

st.title("📄 KLN Freight Invoice → Excel Extractor")
st.caption("Upload KLN invoices → Extract → Download Excel")

uploads = st.file_uploader(
    "Upload KLN PDF files",
    type=["pdf"],
    accept_multiple_files=True,
)

extract_btn = st.button("Extract Invoices", type="primary", disabled=not uploads)

if extract_btn and uploads:
    rows = []
    progress = st.progress(0)
    status = st.empty()

    total = len(uploads)

    for i, f in enumerate(uploads, start=1):
        status.write(f"Parsing: **{f.name}**")
        data = f.read()

        invoice_id = extract_invoice_id(f.name)

        row = parse_invoice_pdf_bytes(data, invoice_id)

        if row:
            rows.append(row)
        else:
            st.warning(f"❌ Could not extract from {f.name}")

        progress.progress(i / total)

    if rows:
        df = pd.DataFrame(rows)

        # Ensure all columns exist
        for col in HEADERS:
            if col not in df.columns:
                df[col] = None

        df = df[HEADERS]

        st.subheader("Preview")
        st.dataframe(df, use_container_width=True)

        # Build Excel
        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.append(HEADERS)

        for _, r in df.iterrows():
            ws.append([r[h] for h in HEADERS])

        wb.save(output)
        output.seek(0)

        st.success(f"Extraction complete: {len(rows)} invoices.")

        st.download_button(
            "⬇️ Download Invoice_Summary.xlsx",
            data=output,
            file_name="Invoice_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
