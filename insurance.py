import streamlit as st
import random
import pdfplumber
import re

st.title("🏢 Insurance Report Generator")

address = st.text_input("📍 Building Address")
currency = st.selectbox("💱 Currency", ["CAD", "USD"])
sqft = st.number_input("📐 Building Square Footage", min_value=0.0, value=0.0)
market_rent = st.number_input(f"💰 Market Rent ({currency} / sqft)", min_value=0.0, value=0.0)

st.markdown("---")
st.subheader("📄 Upload PDF to auto-extract values")
pdf_file = st.file_uploader("Upload Insurance Report PDF", type=["pdf"])

extracted_payroll = None
extracted_rental = None
extracted_turnover = None

DEFAULT_OCR = 0.20  # 20% fallback Occupancy Cost Ratio

def extract_year4_value_flexible(table, key_phrase, target_number_index=4):
    key_phrase = key_phrase.lower()
    for row in table:
        row_text = " ".join(cell.lower() if cell else "" for cell in row)
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
            if cell and phrase in cell.lower():
                if idx + 1 < len(row):
                    val_str = row[idx + 1]
                    if val_str:
                        val_clean = re.sub(r"[^0-9.\-]", "", val_str.replace(',', ''))
                        try:
                            return float(val_clean)
                        except:
                            return None
    return None

if pdf_file is not None:
    with pdfplumber.open(pdf_file) as pdf:
        all_tables = []
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                all_tables.extend(tables)

        for table in all_tables:
            if extracted_payroll is None:
                extracted_payroll = extract_year4_value_flexible(table, "staff costs")
            if extracted_turnover is None:
                extracted_turnover = extract_year4_value_flexible(table, "gross revenue")

        headline_rent = None
        rentable_area = None
        for table in all_tables:
            if headline_rent is None:
                headline_rent = extract_number_next_to_phrase(
                    table,
                    "headline rent (as reviewed by partner) usd psft p.a."
                )
            if rentable_area is None:
                rentable_area = extract_number_next_to_phrase(
                    table,
                    "rentable area sqft"
                )

        if headline_rent is not None and rentable_area is not None:
            extracted_rental = headline_rent * rentable_area
        else:
            extracted_rental = None

if pdf_file:
    st.markdown("### Extracted values from PDF:")
    st.write(f"**Estimated Annual Payroll:** {extracted_payroll if extracted_payroll is not None else 'Not found'}")
    st.write(f"**Rental (Budget/Estimate - Next Year):** {extracted_rental if extracted_rental is not None else 'Not found'}")
    st.write(f"**Annual Turnover (Forecast):** {extracted_turnover if extracted_turnover is not None else 'Not found'}")

if extracted_turnover is not None and extracted_rental is not None:
    gross_profit = extracted_turnover - extracted_rental
else:
    gross_profit = None

if st.button("Generate Report") and address and sqft > 0 and market_rent > 0:

    multi_tenanted = "Yes" if sqft > 10000 else "No"
    building_age = random.randint(20, 50)
    num_floors = max(1, int(sqft // 10000))

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

    # --- RENTAL ESTIMATE FALLBACK ---
    rental_estimate = extracted_rental if extracted_rental is not None else sqft * market_rent

    # --- ANNUAL TURNOVER FALLBACK ---
    if extracted_turnover is not None:
        annual_turnover = extracted_turnover
    elif rental_estimate is not None and rental_estimate > 0:
        annual_turnover = rental_estimate / DEFAULT_OCR
    else:
        annual_turnover = (sqft * market_rent) / DEFAULT_OCR if sqft > 0 and market_rent > 0 else None

    gross_profit_calc = None
    if annual_turnover is not None and rental_estimate is not None:
        gross_profit_calc = annual_turnover - rental_estimate

    st.subheader("🏗️ Building Information")
    st.write(f"**Address:** {address}")
    st.write(f"**Multi-tenanted:** {multi_tenanted}")
    st.write(f"**Approximate Age:** {building_age} years")
    st.write(f"**Total Floors (excl. basement):** {num_floors}")

    st.subheader("👥 Employment Estimate")
    st.write(f"**Estimated FTEs:** {fte}")
    st.write(f"**Estimated Annual Payroll:** {currency} {payroll:,.2f}")

    st.subheader("📊 Forecasted Financials")

    if any([extracted_payroll, extracted_rental, extracted_turnover]):
        st.write(f"**Estimated Annual Payroll (from PDF):** {currency} {extracted_payroll if extracted_payroll is not None else 'N/A':,.2f}")
        st.write(f"**Rental (Budget/Estimate - Next Year) (from PDF or input):** {currency} {rental_estimate:,.2f}")
        st.write(f"**Annual Turnover (Forecast):** {currency} {annual_turnover:,.2f}" if annual_turnover else "Annual Turnover: N/A")
        if gross_profit is not None:
            st.write(f"**Annual Gross Profit (calculated from PDF):** {currency} {gross_profit:,.2f}")
        elif gross_profit_calc is not None:
            st.write(f"**Annual Gross Profit (calculated):** {currency} {gross_profit_calc:,.2f}")
        else:
            st.write("**Annual Gross Profit:** Unable to calculate")
    else:
        st.write(f"**3.5 Rental (Budget/Estimate - Next Year):** {currency} {rental_estimate:,.2f}")
        st.write(f"**3.3 Annual Turnover (Forecast):** {currency} {annual_turnover:,.2f}" if annual_turnover else "Annual Turnover: N/A")
        st.write(f"**3.4 Annual Gross Profit:** {currency} {gross_profit_calc:,.2f}" if gross_profit_calc else "Gross Profit: N/A")

else:
    st.info("Please fill in all fields to generate the report.")
