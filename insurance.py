import streamlit as st
import random

st.title("🏢 Insurance Report Generator")

# User Inputs
address = st.text_input("📍 Building Address")
currency = st.selectbox("💱 Currency", ["CAD", "USD"])
sqft = st.number_input("📐 Building Square Footage", min_value=0.0, value=0.0)
market_rent = st.number_input(f"💰 Market Rent ({currency} / sqft)", min_value=0.0, value=0.0)

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
    gross_profit = rental_estimate - annual_turnover

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
    st.write(f"**3.5 Rental (Budget/Estimate - Next Year):** {currency} {rental_estimate:,.2f}")
    st.write(f"**3.3 Annual Turnover (Forecast):** {currency} {annual_turnover:,.2f}")
    st.write(f"**3.4 Annual Gross Profit:** {currency} {gross_profit:,.2f}")

else:
    st.info("Please fill in all fields to generate the report.")
