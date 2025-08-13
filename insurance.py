import streamlit as st
import random
import pdfplumber
import re
import pandas as pd
import requests
from docx import Document
from docx.shared import RGBColor
from io import BytesIO

st.title("🏢 Insurance Report Generator")

# Hide GitHub icon and hamburger menu
hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header > div:nth-child(1) {visibility: hidden;}
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ---------------- User Inputs ----------------
address = st.text_input(" Building Address")
currency = st.selectbox(" Currency", ["CAD", "USD"])
sqft = st.number_input(" Building Square Footage", min_value=0.0, value=0.0)
market_rent = st.number_input(f" Market Rent ({currency} / sqft)", min_value=0.0, value=0.0)

st.markdown("---")
st.subheader("Optional manual overrides (if OSM data is missing or unavailable)")
multi_tenanted_input = st.selectbox("Is the building multi-tenanted?", ["Unknown", "Yes", "No"], index=0)
building_age_input = st.number_input("Approximate Building Age (years)", min_value=0, max_value=200, value=0)
num_floors_input = st.number_input("Total Floors (excl. basement)", min_value=0, max_value=100, value=0)

st.markdown("---")
st.subheader("📄 Upload Insurance Report PDF file")
pdf_file = st.file_uploader("Upload PDF", type=["pdf"])

DEFAULT_OCR = 0.20  # 20% fallback Occupancy Cost Ratio

# ---------------- Extraction functions ----------------
def extract_value_flexible(table, key_phrase, target_number_index=1):
    key_phrase = key_phrase.lower()
    for row in table:
        row_text = " ".join(str(cell).lower() if cell else "" for cell in row)
        if key_phrase in row_text:
            numbers_str = re.findall(r"[-+]?\d*\.\d+|\d+", row_text.replace(',', ''))
            if len(numbers_str) > target_number_index:
                try:
                    return float(numbers_str[target_number_index])
                except:
                    return None
    return None

def extract_number_next_to_phrase(table, phrase):
    phrase = phrase.lower()
    for row in table:
        for idx, cell in enumerate(row):
            if cell and phrase in str(cell).lower():
                if idx + 1 < len(row):
                    val_str = row[idx + 1]
                    if val_str:
                        val_clean = re.sub(r"[^0-9.\-]", "", str(val_str).replace(',', ''))
                        try:
                            return float(val_clean)
                        except:
                            return None
    return None

def extract_from_pdf(pdf_file):
    payroll = None
    turnover = None
    headline_rent = None
    rentable_area = None
    with pdfplumber.open(pdf_file) as pdf:
        all_tables = []
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                all_tables.extend(tables)

        for table in all_tables:
            if payroll is None:
                raw_payroll = extract_value_flexible(table, "staff costs", target_number_index=3)
                if raw_payroll is not None:
                    payroll = raw_payroll * 1000
            if turnover is None:
                raw_turnover = extract_value_flexible(table, "gross revenue", target_number_index=1)
                if raw_turnover is not None:
                    turnover = raw_turnover * 1000

        for table in all_tables:
            if headline_rent is None:
                headline_rent = extract_number_next_to_phrase(table, "headline rent (as reviewed by partner) usd psft p.a.")
            if rentable_area is None:
                rentable_area = extract_number_next_to_phrase(table, "rentable area sqft")

    rental = None
    if headline_rent is not None and rentable_area is not None:
        rental = headline_rent * rentable_area
    return payroll, rental, turnover

# ---------------- Extract PDF data ----------------
if pdf_file is not None:
    extracted_payroll, extracted_rental, extracted_turnover = extract_from_pdf(pdf_file)
else:
    extracted_payroll, extracted_rental, extracted_turnover = None, None, None

# ---------------- Generate Report Logic ----------------
if st.button("Generate Report") and address and sqft > 0 and market_rent > 0:

    # Multi-tenanted logic: "Unknown" treated as Yes
    multi_tenanted = multi_tenanted_input if multi_tenanted_input != "Unknown" else "Yes"

    # Building age
    building_age = building_age_input if building_age_input > 0 else random.randint(20,50)

    # Floors: manual > fallback
    num_floors = int(num_floors_input) if num_floors_input > 0 else max(1, int(sqft // 10000))

    # Payroll & FTE
    if extracted_payroll is None:
        if sqft < 10000:
            fte = 0.5
            payroll = 50000
        elif sqft < 15000:
            fte = 1.0
            payroll = 65000
        elif sqft < 20000:
            fte = 1.5
            payroll = 110000
        else:
            fte = 2.0
            payroll = 110000
    else:
        payroll = extracted_payroll
        fte = max(0.5, round(payroll/50000,1))  # ensure FTE is always shown

    rental_estimate = extracted_rental if extracted_rental is not None else sqft * market_rent

    if extracted_turnover is not None:
        annual_turnover = extracted_turnover
    else:
        annual_turnover = rental_estimate / DEFAULT_OCR if rental_estimate else (sqft*market_rent)/DEFAULT_OCR

    gross_profit_calc = annual_turnover - rental_estimate if annual_turnover and rental_estimate else 0

    # ---------------- Display in Streamlit ----------------
    st.subheader("🏗 Building Information")
    st.write(f"**Address:** {address}")
    st.write(f"**Multi-tenanted:** {multi_tenanted}")
    st.write(f"**Approximate Age:** {building_age} years")
    st.write(f"**Total Floors (excl. basement):** {num_floors}")

    st.subheader("💼 Employment Estimate")
    st.write(f"**Estimated FTEs:** {fte}")
    st.write(f"**Estimated Annual Payroll:** {currency} {payroll:,.0f}")

    st.subheader("📊 Forecasted Financials")
    st.write(f"**Rental (Budget/Estimate - Next Year):** {currency} {rental_estimate:,.0f}")
    st.write(f"**Annual Turnover (Forecast):** {currency} {annual_turnover:,.0f}")
    st.write(f"**Annual Gross Profit:** {currency} {gross_profit_calc:,.0f}")

    # ---------------- Word Document Generation ----------------
    doc = Document("Insurance Template.docx")
    # Update currency in header
    for para in doc.paragraphs:
        if "All values are in" in para.text:
            para.text = f"Please see the answers in blue below. All values are in {currency} and sq ft."

    # Helper to insert blue text
    def insert_blue_text(para, value):
        para.clear()
        run = para.add_run(str(value))
        run.font.color.rgb = RGBColor(0,0,255)

    for para in doc.paragraphs:
        text = para.text.lower()
        if "multi-tenanted" in text:
            insert_blue_text(para, multi_tenanted)
        elif "approximate age" in text:
            insert_blue_text(para, f"{building_age} years")
        elif "total number of floors" in text:
            insert_blue_text(para, num_floors)
        elif "number of employees" in text:
            insert_blue_text(para, f"{fte} FTE’s")
        elif "3.3 annual turnover" in text:
            continue
        elif "current year forecast" in text and "turnover" in doc.paragraphs[doc.paragraphs.index(para)-1].text.lower():
            insert_blue_text(para, f"${annual_turnover:,.0f}")
        elif "3.4 annual gross profit" in text:
            continue
        elif "current year forecast" in text and "gross profit" in doc.paragraphs[doc.paragraphs.index(para)-1].text.lower():
            insert_blue_text(para, f"${gross_profit_calc:,.0f}")
        elif "3.5 rental" in text:
            continue
        elif "budget/estimate" in text:
            insert_blue_text(para, f"${rental_estimate:,.0f}")
        elif "estimated annual payroll" in text:
            insert_blue_text(para, f"${payroll:,.0f}")

    # Streamlit download
    doc_stream = BytesIO()
    doc.save(doc_stream)
    doc_stream.seek(0)
    st.download_button(
        label="📄 Download Word Report",
        data=doc_stream,
        file_name=f"Insurance_Report_{address.replace(' ','_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

else:
    st.info("Please fill in all fields to generate the report.")
