import streamlit as st
import random
import pdfplumber
import re

st.title("🏢 Insurance Report Generator")

# User Inputs
address = st.text_input("📍 Building Address")
currency = st.selectbox("💱 Currency", ["CAD", "USD"])
sqft = st.number_input("📐 Building Square Footage", min_value=0.0, value=0.0)
market_rent = st.number_input(f"💰 Market Rent ({currency} / sqft)", min_value=0.0, value=0.0)

st.markdown("---")
st.subheader("📄 Upload PDF to auto-extract values")
pdf_file = st.file_uploader("Upload Insurance Report PDF", type=["pdf"])

# Variables to hold extracted values from PDF if uploaded
extracted_payroll = None
extracted_rental = None
extracted_turnover = None

def extract_year4_value_flexible(table, key_phrase, target_number_index=4):
    """
    For each row in the table, if the key_phrase is found anywhere in the row (any cell),
    extract all numbers from that row, then pick the number at target_number_index (0-based).
    If no suitable number found, return None.
    """
    key_phrase = key_phrase.lower()
    for row in table:
        # Join all cells text and search for the key phrase
        row_text = " ".join(cell.lower() if cell else "" for cell in row)
        if key_phrase in row_text:
            # Extract all numbers from the row text
            numbers_str = re.findall(r"[-+]?\d*\.\d+|\d+", row_text.replace(',', ''))
            if len(numbers_str) > target_number_index:
                try:
                    return float(numbers_str[target_number_index])
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
            if extracted_rental is None:
                extracted_rental = extract_year4_value_flexible(table, "market rent")
            if extracted_turnover is None:
                extracted_turnover = extract_year4_value_flexible(table, "gross revenue")

# Show extracted values if found
if pdf_file:
    st.markdown("### Extracted values from PDF:")
    st.write(f"**Estimated Annual Payroll:** {extracted_payroll if extracted_payroll is not None else 'Not found'}")
    st.write(f"**Rental (Budget/Estimate - Next Year):** {extracted_rental if extracted_rental is not None else 'Not found'}")
    st.write(f"**Annual Turnover (Forecast):** {extracted_turnover if extracted_turnover is not None else 'Not found'}")

# Calculate gross profit from extracted PDF values if available
if extracted_turnover is not None and extracted_rental is not None:
    gross_profit = extracted_turnover - extracted_rental
else:
    gross_profit = None

# Button to generate report
if st.button("Generate Report") and address and sqft > 0 and market_rent > 0:

    # Inferred/Mocked Data
    multi_tenanted = "Yes" if sqft > 10000 else "No"
    building_age = random.randint(20, 50)  # years
    num_floors = max(1, int(sqft // 10000))  # 1 floor per 10k sqft

    # FTE Logic
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

    # Financial Calculations
    rental_estimate = sqft * market_rent
    annual_turnover = rental_estimate * 2
    gross_profit_calc = annual_turnover - rental_estimate  # corrected from original (turnover - rental)

    # Report Display
    st.subheader("🏗️ Building Information")
    st.write(f"**Address:** {address}")
    st.write(f"**Multi-tenanted:** {multi_tenanted}")
    st.write(f"**Approximate Age:** {building_age} years")
    st.write(f"**Total Floors (excl. basement):** {num_floors}")

    st.subheader("👥 Employment Estimate")
    st.write(f"**Estimated FTEs:** {fte}")
    st.write(f"**Estimated Annual Payroll:** {currency} {payroll:,.2f}")

    st.subheader("📊 Forecasted Financials")

    # If PDF extraction successful, show those values, else use original calculations
    if any([extracted_payroll, extracted_rental, extracted_turnover]):
        st.write(f"**Estimated Annual Payroll (from PDF):** {currency} {extracted_payroll if extracted_payroll is not None else 'N/A':,.2f}")
        st.write(f"**Rental (Budget/Estimate - Next Year) (from PDF):** {currency} {extracted_rental if extracted_rental is not None else 'N/A':,.2f}")
        st.write(f"**Annual Turnover (Forecast) (from PDF):** {currency} {extracted_turnover if extracted_turnover is not None else 'N/A':,.2f}")
        if gross_profit is not None:
            st.write(f"**Annual Gross Profit (calculated from PDF):** {currency} {gross_profit:,.2f}")
        else:
            st.write("**Annual Gross Profit:** Unable to calculate from PDF data")
    else:
        st.write(f"**3.5 Rental (Budget/Estimate - Next Year):** {currency} {rental_estimate:,.2f}")
        st.write(f"**3.3 Annual Turnover (Forecast):** {currency} {annual_turnover:,.2f}")
        st.write(f"**3.4 Annual Gross Profit:** {currency} {gross_profit_calc:,.2f}")

else:
    st.info("Please fill in all fields to generate the report.")
