import streamlit as st
import random
import pdfplumber
import re
import pandas as pd

st.title("🏢 Insurance Report Generator")

address = st.text_input("📍 Building Address")
currency = st.selectbox("💱 Currency", ["CAD", "USD"])
sqft = st.number_input("📐 Building Square Footage", min_value=0.0, value=0.0)
market_rent = st.number_input(f"💰 Market Rent ({currency} / sqft)", min_value=0.0, value=0.0)

st.markdown("---")
st.subheader("📄 Upload Insurance Report PDF or Excel file")

pdf_file = st.file_uploader("Upload PDF", type=["pdf"])
excel_file = st.file_uploader("Upload Excel Workbook", type=["xlsx"])

extracted_payroll = None
extracted_rental = None
extracted_turnover = None

DEFAULT_OCR = 0.20  # 20% fallback Occupancy Cost Ratio

def extract_value_flexible(table, key_phrase, target_number_index=1):
    """
    Extracts number from row containing key_phrase at target_number_index column (0-based).
    """
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

def extract_staff_cost_below_row(table, key_phrase, target_number_index=3):
    """
    Finds the row containing key_phrase (case insensitive),
    then extracts the number at target_number_index from the **next row** (row+1),
    returns None if not found or next row does not exist.
    """
    key_phrase = key_phrase.lower()
    for i, row in enumerate(table):
        row_text = " ".join(str(cell).lower() if cell else "" for cell in row)
        if key_phrase in row_text:
            # check if next row exists
            if i + 1 < len(table):
                next_row = table[i + 1]
                if len(next_row) > target_number_index:
                    val_str = next_row[target_number_index]
                    if val_str:
                        val_clean = re.sub(r"[^0-9.\-]", "", str(val_str).replace(',', ''))
                        try:
                            return float(val_clean)
                        except:
                            return None
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

        # Extract payroll from the row BELOW "staff costs", year 3 (index 3)
        for table in all_tables:
            if payroll is None:
                raw_payroll = extract_staff_cost_below_row(table, "staff costs", target_number_index=3)  # Year 3 col
                if raw_payroll is not None:
                    payroll = raw_payroll * 1000  # multiply by 1000
            if turnover is None:
                raw_turnover = extract_value_flexible(table, "gross revenue", target_number_index=1)  # Year 1 col
                if raw_turnover is not None:
                    turnover = raw_turnover * 1000  # multiply by 1000

        for table in all_tables:
            if headline_rent is None:
                headline_rent = extract_number_next_to_phrase(table, "headline rent (as reviewed by partner) usd psft p.a.")
            if rentable_area is None:
                rentable_area = extract_number_next_to_phrase(table, "rentable area sqft")

    rental = None
    if headline_rent is not None and rentable_area is not None:
        rental = headline_rent * rentable_area
    return payroll, rental, turnover

def extract_from_excel(excel_file):
    payroll = None
    rental = None
    turnover = None
    try:
        xls = pd.ExcelFile(excel_file)
        if "Owned Summary" in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name="Owned Summary")
            df.columns = df.columns.astype(str)
            df_lower = df.applymap(lambda x: str(x).lower() if isinstance(x, str) else x)

            # Staff costs from Year 1 (index 1)
            staff_row_idx = df_lower.index[df_lower.apply(lambda row: row.astype(str).str.contains("staff costs").any(), axis=1)]
            if len(staff_row_idx) > 0:
                try:
                    payroll = float(df.iloc[staff_row_idx[0], 1])
                except:
                    payroll = None

            # Gross revenue from Year 1 (index 1)
            gross_rev_idx = df_lower.index[df_lower.apply(lambda row: row.astype(str).str.contains("gross revenue").any(), axis=1)]
            if len(gross_rev_idx) > 0:
                try:
                    turnover = float(df.iloc[gross_rev_idx[0], 1])
                except:
                    turnover = None

            headline_rent = None
            rentable_area = None
            for idx, row in df_lower.iterrows():
                for col_idx, val in enumerate(row):
                    if isinstance(val, str) and "headline rent (as reviewed by partner)" in val:
                        try:
                            headline_rent = float(df.iloc[idx, col_idx + 1])
                        except:
                            headline_rent = None
                    if isinstance(val, str) and "rentable area sqft" in val:
                        try:
                            rentable_area = float(df.iloc[idx, col_idx + 1])
                        except:
                            rentable_area = None
            if headline_rent is not None and rentable_area is not None:
                rental = headline_rent * rentable_area

    except Exception as e:
        st.error(f"Error reading Excel file: {e}")

    return payroll, rental, turnover

if excel_file is not None:
    extracted_payroll, extracted_rental, extracted_turnover = extract_from_excel(excel_file)
elif pdf_file is not None:
    extracted_payroll, extracted_rental, extracted_turnover = extract_from_pdf(pdf_file)
else:
    extracted_payroll, extracted_rental, extracted_turnover = None, None, None

if excel_file or pdf_file:
    st.markdown("### Extracted values from uploaded file:")
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

    # Use embedded payroll only if no extracted payroll found
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
        fte = None  # Optionally you could estimate FTE from payroll if you want

    rental_estimate = extracted_rental if extracted_rental is not None else sqft * market_rent

    # Annual turnover with fallback to OCR logic
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
    if extracted_payroll is not None:
        st.write(f"**Estimated Annual Payroll (from file):** {currency} {payroll:,.2f}")
        if fte is not None:
            st.write(f"**Estimated FTEs:** {fte}")
    else:
        st.write(f"**Estimated FTEs:** {fte}")
        st.write(f"**Estimated Annual Payroll:** {currency} {payroll:,.2f}")

    st.subheader("📊 Forecasted Financials")
    if any([extracted_payroll, extracted_rental, extracted_turnover]):
        st.write(f"**Rental (Budget/Estimate - Next Year) (from file or input):** {currency} {rental_estimate:,.2f}")
        st.write(f"**Annual Turnover (Forecast):** {currency} {annual_turnover:,.2f}" if annual_turnover else "Annual Turnover: N/A")
        if gross_profit is not None:
            st.write(f"**Annual Gross Profit (calculated from file):** {currency} {gross_profit:,.2f}")
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
