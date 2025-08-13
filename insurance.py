import streamlit as st
import random
import pdfplumber
import re
import pandas as pd
import requests
from io import BytesIO
from docx import Document

st.title("🏢 Insurance Report Generator")

# Hide Streamlit UI elements
hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header > div:nth-child(1) {visibility: hidden;}
    </style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# --- User inputs ---
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

DEFAULT_OCR = 0.20  # fallback Occupancy Cost Ratio

# --- Extraction functions ---
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

# --- Extract PDF data if uploaded ---
if pdf_file is not None:
    extracted_payroll, extracted_rental, extracted_turnover = extract_from_pdf(pdf_file)
else:
    extracted_payroll, extracted_rental, extracted_turnover = None, None, None

if pdf_file:
    st.markdown("### Extracted values from uploaded file:")
    st.write(f"**Estimated Annual Payroll:** {extracted_payroll if extracted_payroll is not None else 'Not found'}")
    st.write(f"**Rental (Budget/Estimate - Next Year):** {extracted_rental if extracted_rental is not None else 'Not found'}")
    st.write(f"**Annual Turnover (Forecast):** {extracted_turnover if extracted_turnover is not None else 'Not found'}")

if extracted_turnover is not None and extracted_rental is not None:
    gross_profit = extracted_turnover - extracted_rental
else:
    gross_profit = None

# --- OSM Geocoding & Overpass API ---
def geocode_address(address):
    url = "https://nominatim.openstreetmap.org/search"
    params = {"q": address, "format": "json", "limit": 1}
    try:
        resp = requests.get(url, params=params, timeout=10, headers={"User-Agent": "StreamlitApp"})
        resp.raise_for_status()
        data = resp.json()
        if data:
            return float(data[0]["lat"]), float(data[0]["lon"])
    except:
        return None, None
    return None, None

def get_building_floors_osm(lat, lon):
    if lat is None or lon is None:
        return None
    query = f"""
    [out:json];
    (
      way(around:30,{lat},{lon})["building"];
      relation(around:30,{lat},{lon})["building"];
      node(around:30,{lat},{lon})["building"];
    );
    out tags 1;
    """
    url = "https://overpass-api.de/api/interpreter"
    try:
        resp = requests.post(url, data=query, timeout=15)
        resp.raise_for_status()
        data = resp.json()
        elements = data.get("elements", [])
        for el in elements:
            tags = el.get("tags", {})
            if "building:levels" in tags:
                floors_str = tags["building:levels"]
                floors = int(re.findall(r'\d+', floors_str)[0])
                return floors
    except:
        return None
    return None

# --- Generate report ---
if st.button("Generate Report") and address and sqft > 0 and market_rent > 0:

    # Get floors from OSM
    lat, lon = geocode_address(address)
    osm_floors = get_building_floors_osm(lat, lon)

    # Multi-tenanted logic: Unknown defaults to Yes
    multi_tenanted = multi_tenanted_input if multi_tenanted_input != "Unknown" else "Yes"

    # Building age logic
    building_age = building_age_input if building_age_input > 0 else random.randint(20, 50)

    # Floors logic
    if num_floors_input > 0:
        num_floors = int(num_floors_input)
    elif osm_floors:
        num_floors = osm_floors
    else:
        num_floors = max(1, int(sqft // 10000))

    # Payroll & FTE logic
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
        fte = 1.0  # Always show FTE even if payroll is from file

    rental_estimate = extracted_rental if extracted_rental is not None else sqft * market_rent

    if extracted_turnover is not None:
        annual_turnover = extracted_turnover
    elif rental_estimate is not None and rental_estimate > 0:
        annual_turnover = rental_estimate / DEFAULT_OCR
    else:
        annual_turnover = (sqft * market_rent) / DEFAULT_OCR if sqft > 0 and market_rent > 0 else None

    gross_profit_calc = None
    if annual_turnover is not None and rental_estimate is not None:
        gross_profit_calc = annual_turnover - rental_estimate

    # --- Display UI ---
    st.subheader("🏗 Building Information")
    st.write(f"**Address:** {address}")
    st.write(f"**Multi-tenanted:** {multi_tenanted}")
    st.write(f"**Approximate Age:** {building_age} years")
    st.write(f"**Total Floors (excl. basement):** {num_floors}")

    st.subheader(" Employment Estimate")
    st.write(f"**Estimated FTEs:** {fte}")
    st.write(f"**Estimated Annual Payroll:** {currency} {payroll:,.2f}")

    st.subheader(" Forecasted Financials")
    st.write(f"**Rental (Budget/Estimate - Next Year):** {currency} {rental_estimate:,.2f}")
    st.write(f"**Annual Turnover (Forecast):** {currency} {annual_turnover:,.2f}" if annual_turnover else "Annual Turnover: N/A")
    st.write(f"**Annual Gross Profit:** {currency} {gross_profit_calc:,.2f}" if gross_profit_calc else "Gross Profit: N/A")

    # --- Generate Word Report ---
    try:
        doc = Document("Insurance Template.docx")

        # Update header currency
        for para in doc.paragraphs:
            if "All values are in" in para.text:
                para.text = f"Please see the answers in blue below. All values are in {currency} and sq ft."

        # Table $ placeholders mapping
        table_mapping = {
            "$AnnualTurnover": f"{annual_turnover:,.2f}" if annual_turnover else "N/A",
            "$GrossProfit": f"{gross_profit_calc:,.2f}" if gross_profit_calc else "N/A",
            "$Rental": f"{rental_estimate:,.2f}",
            "$Payroll": f"{payroll:,.2f}",
        }

        # Replace $ placeholders in table
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for key, val in table_mapping.items():
                        if key in cell.text:
                            cell.text = cell.text.replace(key, str(val))

        # Add FTE after question
        for para in doc.paragraphs:
            if "Number of employees will be employed" in para.text:
                para.add_run(f" {fte}")

        # Save to memory
        doc_stream = BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)

        st.download_button(
            label="📄 Download Word Report",
            data=doc_stream,
            file_name=f"Insurance_Report_{address.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(f"Error generating Word report: {e}")

else:
    st.info("Please fill in all fields to generate the report.")
