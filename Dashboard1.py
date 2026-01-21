import streamlit as st
import pandas as pd
import plotly.express as px

# -----------------------------
# Page Configuration
# -----------------------------
st.set_page_config(
    page_title="Sales Executive Performance Dashboard",
    layout="wide"
)

# -----------------------------
# Indian Currency Formatter
# -----------------------------
def format_inr(amount):
    try:
        amount = float(amount)
    except:
        return "₹0.00"

    s = f"{amount:,.2f}"
    parts = s.split(".")
    integer = parts[0].replace(",", "")

    if len(integer) > 3:
        last3 = integer[-3:]
        rest = integer[:-3]
        rest = ",".join([rest[max(i-2, 0):i] for i in range(len(rest), 0, -2)][::-1])
        return f"₹{rest},{last3}.{parts[1]}"
    else:
        return f"₹{integer}.{parts[1]}"

# -----------------------------
# Smart File Cleaner (ALL formats)
# -----------------------------
@st.cache_data
def clean_uploaded_data(uploaded_file):

    df = pd.read_excel(uploaded_file)

    # Remove empty columns
    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]

    # Normalize column names
    df.columns = df.columns.astype(str).str.strip().str.lower()

    column_map = {}

    for col in df.columns:
        if "sales" in col or "executive" in col or "name" in col:
            column_map[col] = "Sales Person Name"
        elif "night" in col:
            column_map[col] = "Nights"
        elif "occupancy" in col:
            column_map[col] = "Occupancy %"
        elif "pax" in col or "guest" in col:
            column_map[col] = "Pax"
        elif "revenue" in col:
            column_map[col] = "Room Revenue"
        elif "arr" in col:
            column_map[col] = "ARR"
        elif "arp" in col:
            column_map[col] = "ARP"
        elif "%" in col:
            column_map[col] = "Revenue%"

    df = df.rename(columns=column_map)

    # Mandatory columns
    required = ["Sales Person Name", "Room Revenue"]

    for col in required:
        if col not in df.columns:
            raise ValueError(f"Required column missing: {col}")

    # Optional columns – create if missing
    defaults = {
        "Nights": 0,
        "Occupancy %": 0,
        "Pax": 0,
        "Revenue%": 0,
        "ARR": 0,
        "ARP": 0
    }

    for col, val in defaults.items():
        if col not in df.columns:
            df[col] = val

    # Clean numeric columns
    numeric_cols = ["Nights", "Occupancy %", "Pax", "Room Revenue", "Revenue%", "ARR", "ARP"]

    for col in numeric_cols:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace("₹", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.strip()
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Remove totals
    df = df[~df["Sales Person Name"].astype(str).str.contains("total|grand", case=False, na=False)]

    return df

# -----------------------------
# KPI Styling
# -----------------------------
st.markdown("""
<style>
.kpi-box {
    padding: 20px;
    border-radius: 12px;
    background-color: #f6f8fa;
    text-align: center;
    box-shadow: 0 2px 6px rgba(0,0,0,0.05);
}
.kpi-title { font-size: 15px; color: #666; }
.kpi-value { font-size: 24px; font-weight: 700; white-space: nowrap; }
</style>
""", unsafe_allow_html=True)

# -----------------------------
# UI
# -----------------------------
st.title("Sales Executive Performance Dashboard")

uploaded_file = st.file_uploader(
    "Upload Sales Executive Excel File (.xls or .xlsx)",
    type=["xls", "xlsx"]
)

if uploaded_file:

    try:
        df = clean_uploaded_data(uploaded_file)

        # Sidebar
        st.sidebar.header("Dashboard Controls")

        selected_names = st.sidebar.multiselect(
            "Select Sales Persons",
            df["Sales Person Name"].unique(),
            default=df["Sales Person Name"].unique()
        )

        filtered_df = df[df["Sales Person Name"].isin(selected_names)]

        # KPIs
        total_revenue = filtered_df["Room Revenue"].sum()
        total_nights = int(filtered_df["Nights"].sum())
        avg_arr = filtered_df["ARR"].mean()
        total_pax = int(filtered_df["Pax"].sum())

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            st.markdown(f"<div class='kpi-box'><div class='kpi-title'>Total Revenue</div><div class='kpi-value'>{format_inr(total_revenue)}</div></div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div class='kpi-box'><div class='kpi-title'>Total Nights</div><div class='kpi-value'>{total_nights}</div></div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"<div class='kpi-box'><div class='kpi-title'>Average ARR</div><div class='kpi-value'>{format_inr(avg_arr)}</div></div>", unsafe_allow_html=True)
        with c4:
            st.markdown(f"<div class='kpi-box'><div class='kpi-title'>Total Pax</div><div class='kpi-value'>{total_pax}</div></div>", unsafe_allow_html=True)

        st.divider()

        # Charts
        col1, col2 = st.columns(2)

        with col1:
            fig_bar = px.bar(
                filtered_df.sort_values("Room Revenue"),
                x="Room Revenue",
                y="Sales Person Name",
                orientation="h"
            )
            fig_bar.update_layout(xaxis_tickprefix="₹ ")
            st.plotly_chart(fig_bar, use_container_width=True)

        with col2:
            fig_pie = px.pie(
                filtered_df,
                values="Room Revenue",
                names="Sales Person Name",
                hole=0.4
            )
            st.plotly_chart(fig_pie, use_container_width=True)

        st.divider()

        fig_scatter = px.scatter(
            filtered_df,
            x="Nights",
            y="Room Revenue",
            size="Room Revenue",
            color="Sales Person Name"
        )
        fig_scatter.update_layout(yaxis_tickprefix="₹ ")
        st.plotly_chart(fig_scatter, use_container_width=True)

        with st.expander("View Data"):
            st.dataframe(filtered_df, use_container_width=True)

    except Exception as e:
        st.error(f"Error processing file: {e}")

else:
    st.info("Please upload a Sales Executive Excel file to generate the dashboard.")
