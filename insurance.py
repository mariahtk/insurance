import streamlit as st
import requests
import random

# Geocode address to lat/lon using free Nominatim API
def geocode_address(address):
    url = "https://nominatim.openstreetmap.org/search"
    params = {"q": address, "format": "json"}
    try:
        response = requests.get(url, params=params, headers={"User-Agent": "insurance-report-app"})
        results = response.json()
        if results:
            return float(results[0]["lat"]), float(results[0]["lon"])
    except Exception as e:
        st.error(f"Geocoding error: {e}")
    return None, None

# Query Overpass API to get building floors (building:levels tag)
def get_floors_from_osm(lat, lon):
    query = f"""
    [out:json];
    (
      way(around:50,{lat},{lon})["building"];
      relation(around:50,{lat},{lon})["building"];
    );
    out tags 1;
    """
    try:
        response = requests.post("https://overpass-api.de/api/interpreter", data={"data": query})
        data = response.json()
        for element in data.get("elements", []):
            tags = element.get("tags", {})
            if "building:levels" in tags:
                try:
                    return int(tags["building:levels"])
                except:
                    pass
    except Exception as e:
        st.error(f"OSM Overpass API error: {e}")
    return None

st.title("🏢 Insurance Report Generator with OSM Floors Lookup")

# User inputs
address = st.text_input("📍 Building Address")
currency = st.selectbox("💱 Currency", ["CAD", "USD"])
sqft = st.number_input("📐 Building Square Footage", min_value=0.0, value=0.0)
market_rent = st.number_input(f"💰 Market Rent ({currency} / sqft)", min_value=0.0, value=0.0)

if st.button("Generate Report"):

    if not address or sqft <= 0 or market_rent <= 0:
        st.error("Please fill in all fields with valid values.")
    else:
        # Geocode
        lat, lon = geocode_address(address)
        if lat is None or lon is None:
            st.warning("Could not geocode address. Using fallback floor estimate.")
            num_floors = max(1, int(sqft // 10000))
        else:
            # Get floors from OSM
            num_floors = get_floors_from_osm(lat, lon)
            if num_floors is None:
                st.info("Building floor data not found in OSM. Using fallback estimate.")
                num_floors = max(1, int(sqft // 10000))

        # Multi-tenanted logic (simple heuristic)
        multi_tenanted = "Yes" if sqft > 10000 else "No"

        # Approximate age (random for demo)
        building_age = random.randint(20, 50)

        # FTE and Payroll logic
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

        # Financials
        rental_estimate = sqft * market_rent
        annual_turnover = rental_estimate * 2
        gross_profit = rental_estimate - annual_turnover

        # Display output
        st.subheader("🏗️ Building Information")
        st.write(f"**Address:** {address}")
        st.write(f"**Latitude, Longitude:** {lat}, {lon}")
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
