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

def extract_year4_value(table, row_label):
    """
    Given a table (list of lists), find the row where first column matches row_label
    and return the value under the Year 4 column (assumed to be the 5th column, index 4).
    """
    for row in table:
        if len(row) > 4 and row[0].strip().lower() == row_label.strip().lower():
            # Clean the value to extract numbers only, remove commas and currency symbols
            val_str = row[4]
            if val_str:
                val_clean = re.sub(r"[^0-9.\-]", "", val_str)
                try:
                    return float(val_clean)
                except:
                    return None
    return None

if pdf_file is not None:
    with pdfplumber.open(pdf_file) as pdf:
        # Aggregate tables from all pages to search
        all_tables = []
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                all_tables.extend(tables)
        
        # Try to extract the values based on the row labels you gave
        for table in all_tables:
            if extracted_payroll is None:
                extracted_payroll = extract_year4_value(table, "Staff Costs")
            if extracted_rental is None:
                extracted_rental = extract_year4_value(table, "Market Rent (as reviewed by partner)")
            if extracted_turnover is None:
                extracted_turnover = extract_year4_value(table, "Gross Revenue")

# Show extracted values if found
if pdf_file:
    st.markdown("### Extracted values from PDF:")
    st.write(f"**Estimated Annual Payroll:** {extracted_payroll if extracted_payroll is not None else 'Not found'}")
    st.write(f"**Rental (Budget/Estimate - Next Year):** {extracted_rental if extracted_rental is not None else 'Not found'}")
    st.write(f"**Annual Turnover (Forecast):** {extracted_turnover if extracted_turnover is not None else 'Not found'}")

# If PDF extracted all needed values, calculate gross profit from those
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
    gross_profit_calc = rental_estimate - annual_turnover

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
    if extracted_payroll is not None or extracted_rental is not None or extracted_turnover is not None:
        st.write(f"**Estimated Annual Payroll (from PDF):** {currency} {extracted_payroll if extracted_payroll else 'N/A':,.2f}")
        st.write(f"**Rental (Budget/Estimate - Next Year) (from PDF):** {currency} {extracted_rental if extracted_rental else 'N/A':,.2f}")
        st.write(f"**Annual Turnover (Forecast) (from PDF):** {currency} {extracted_turnover if extracted_turnover else 'N/A':,.2f}")
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
