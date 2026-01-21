import streamlit as st
import pandas as pd
import plotly.express as px
import re

# -----------------------------
# Page Configuration
# -----------------------------
st.set_page_config(page_title="Sales Executive Performance Dashboard", layout="wide")

# -----------------------------
# INR Formatter
# -----------------------------
def format_inr(amount):
    try:
        amount = float(amount)
    except:
        return "₹0.00"
    return f"₹{amount:,.2f}"

# -----------------------------
# Column Auto Detection
# -----------------------------
def find_column(columns, keywords):
    for col in columns:
        name = col.lower()
        for k in keywords:
            if k in name:
                return col
    return None

# -----------------------------
# Clean Uploaded File
# -----------------------------
@st.cache_data
def clean_uploaded_data(uploaded_file):

    df = pd.read_excel(uploaded_file)

    # Drop empty columns
    df = df.dropna(axis=1, how="all")

    # Normalize column names
    df.columns = [str(c).strip() for c in df.columns]

    cols = df.columns.tolist()

    # Auto-detect columns
    col_name = find_column(cols, ["sales", "executive", "name", "person"])
    col_nights = find_column(cols, ["night"])
    col_pax = find_column(cols, ["pax", "guest"])
    col_revenue = find_column(cols, ["revenue", "amount", "room rev", "sales"])
    col_arr = find_column(cols, ["arr", "rate"])
    col_occ = find_column(cols, ["occupancy", "occ"])
    col_rev_pct = find_column(cols, ["revenue%", "rev%"])
    col_arp = find_column(cols, ["arp"])

    if not col_name or not col_revenue:
        raise ValueError("Required columns not found in Excel file (Name or Revenue).")

    # Build standardized dataframe
    new_df = pd.DataFrame()
    new_df["Sales Person Name"] = df[col_name]
    new_df["Nights"] = df[col_nights] if col_nights else 0
    new_df["Pax"] = df[col_pax] if col_pax else 0
    new_df["Room Revenue"] = df[col_revenue]
    new_df["ARR"] = df[col_arr] if col_arr else 0
    new_df["Occupancy %"] = df[col_occ] if col_occ else 0
    new_df["Revenue%"] = df[col_rev_pct] if col_rev_pct else 0
    new_df["ARP"] = df[col_arp] if col_arp else 0

    # Remove totals
    new_df = new_df.dropna(subset=["Sales Person Name"])
    new_df = new_df[~new_df["Sales Person Name"].astype(str).str.contains("total", case=False)]

    # Convert numeric columns
    for c in ["Nights", "Pax", "Room Revenue", "ARR", "Occupancy %", "Revenue%", "ARP"]:
        new_df[c] = (
            new_df[c]
            .astype(str)
            .str.replace("₹", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.strip()
        )
        new_df[c] = pd.to_numeric(new_df[c], errors="coerce").fillna(0)

    return new_df

# -----------------------------
# KPI CSS
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
.kpi-value { font-size: 24px; font-weight: 700; }
</style>
""", unsafe_allow_html=True)

# -----------------------------
# UI
# -----------------------------
st.title("Sales Executive Performance Dashboard")

uploaded_file = st.file_uploader("Upload Excel File", type=["xls", "xlsx"])

if uploaded_file:

    try:
        df = clean_uploaded_data(uploaded_file)

        st.sidebar.header("Filters")
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

        for col, title, value in [
            (c1, "Total Room Revenue", format_inr(total_revenue)),
            (c2, "Total Nights Sold", total_nights),
            (c3, "Average Room Rate (ARR)", format_inr(avg_arr)),
            (c4, "Total Pax", total_pax),
        ]:
            col.markdown(f"""
            <div class="kpi-box">
                <div class="kpi-title">{title}</div>
                <div class="kpi-value">{value}</div>
            </div>
            """, unsafe_allow_html=True)

        st.divider()

        # Bar chart
        fig_bar = px.bar(
            filtered_df.sort_values("Room Revenue"),
            x="Room Revenue",
            y="Sales Person Name",
            orientation="h"
        )
        fig_bar.update_layout(xaxis_tickprefix="₹ ")
        st.plotly_chart(fig_bar, use_container_width=True)

        # Pie chart
        fig_pie = px.pie(filtered_df, values="Room Revenue", names="Sales Person Name", hole=0.4)
        st.plotly_chart(fig_pie, use_container_width=True)

        # Scatter
        fig_scatter = px.scatter(
            filtered_df, x="Nights", y="Room Revenue", size="Room Revenue",
            color="Sales Person Name"
        )
        fig_scatter.update_layout(yaxis_tickprefix="₹ ")
        st.plotly_chart(fig_scatter, use_container_width=True)

        with st.expander("View Data"):
            st.dataframe(filtered_df, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")

else:
    st.info("Upload your Excel file to start.")
