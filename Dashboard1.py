import streamlit as st
import pandas as pd
import plotly.express as px
import os

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
# Smart Column Detector
# -----------------------------
def find_col(df, keywords):
    for col in df.columns:
        for kw in keywords:
            if kw.lower() in str(col).lower():
                return col
    return None

# -----------------------------
# Clean Uploaded File (ROBUST)
# -----------------------------
@st.cache_data
def clean_uploaded_data(uploaded_file):

    # Try reading with different skip rows
    for skip in [0, 1, 2, 3, 4]:
        try:
            df = pd.read_excel(uploaded_file, skiprows=skip)
            if len(df.columns) >= 5:
                break
        except:
            continue

    # Remove unnamed columns
    df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]

    df.columns = [str(c).strip() for c in df.columns]

    col_sales = find_col(df, ["sales", "executive", "person", "name"])
    col_nights = find_col(df, ["night"])
    col_occ = find_col(df, ["occupancy"])
    col_pax = find_col(df, ["pax", "guest"])
    col_revenue = find_col(df, ["room", "revenue", "rev"])
    col_arr = find_col(df, ["arr"])
    col_arp = find_col(df, ["arp"])
    col_revperc = find_col(df, ["revenue%", "rev%", "contribution"])

    if col_revenue is None:
        raise ValueError("Room Revenue column not detected automatically.")

    rename_map = {
        col_sales: "Sales Person Name",
        col_nights: "Nights",
        col_occ: "Occupancy %",
        col_pax: "Pax",
        col_revenue: "Room Revenue",
        col_arr: "ARR",
        col_arp: "ARP",
        col_revperc: "Revenue%"
    }

    rename_map = {k: v for k, v in rename_map.items() if k}
    df = df.rename(columns=rename_map)

    if "Revenue%" not in df.columns:
        df["Revenue%"] = 0

    df = df.dropna(subset=["Sales Person Name"])
    df = df[~df["Sales Person Name"].astype(str).str.contains(
        "total|grand|undefined", case=False, na=False
    )]

    numeric_cols = ["Nights", "Occupancy %", "Pax", "Room Revenue", "Revenue%", "ARR", "ARP"]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = (
                df[col].astype(str)
                .str.replace("₹", "", regex=False)
                .str.replace(",", "", regex=False)
                .str.strip()
            )
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

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
.kpi-title {
    font-size: 15px;
    color: #666;
}
.kpi-value {
    font-size: 24px;
    font-weight: 700;
}
</style>
""", unsafe_allow_html=True)

# -----------------------------
# App UI
# -----------------------------
st.title("Sales Executive Performance Dashboard")

uploaded_file = st.file_uploader(
    "Upload Sales Executive Excel File (.xls or .xlsx)",
    type=["xls", "xlsx"]
)

if uploaded_file:

    try:
        df = clean_uploaded_data(uploaded_file)

        st.sidebar.header("Dashboard Controls")
        selected_names = st.sidebar.multiselect(
            "Select Sales Persons",
            options=df["Sales Person Name"].unique(),
            default=df["Sales Person Name"].unique()
        )

        filtered_df = df[df["Sales Person Name"].isin(selected_names)]

        st.subheader("Performance Analysis")
        st.divider()

        total_revenue = filtered_df["Room Revenue"].sum()
        total_nights = int(filtered_df["Nights"].sum())
        avg_arr = filtered_df["ARR"].mean()
        total_pax = int(filtered_df["Pax"].sum())

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            st.markdown(f"""<div class="kpi-box"><div class="kpi-title">Total Room Revenue</div><div class="kpi-value">{format_inr(total_revenue)}</div></div>""", unsafe_allow_html=True)

        with c2:
            st.markdown(f"""<div class="kpi-box"><div class="kpi-title">Total Nights Sold</div><div class="kpi-value">{total_nights}</div></div>""", unsafe_allow_html=True)

        with c3:
            st.markdown(f"""<div class="kpi-box"><div class="kpi-title">Average Room Rate (ARR)</div><div class="kpi-value">{format_inr(avg_arr)}</div></div>""", unsafe_allow_html=True)

        with c4:
            st.markdown(f"""<div class="kpi-box"><div class="kpi-title">Total Pax (Guests)</div><div class="kpi-value">{total_pax}</div></div>""", unsafe_allow_html=True)

        st.divider()

        col_left, col_right = st.columns(2)

        with col_left:
            fig_rev = px.bar(
                filtered_df.sort_values("Room Revenue"),
                x="Room Revenue",
                y="Sales Person Name",
                orientation="h",
                color="Room Revenue"
            )
            fig_rev.update_layout(xaxis=dict(tickprefix="₹ ", tickformat=",.0f"))
            st.plotly_chart(fig_rev, use_container_width=True)

        with col_right:
            fig_pie = px.pie(filtered_df, values="Room Revenue", names="Sales Person Name", hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)

        st.subheader("Efficiency Analysis: Nights vs Revenue")

        fig_scatter = px.scatter(
            filtered_df,
            x="Nights",
            y="Room Revenue",
            size="Room Revenue",
            color="Sales Person Name"
        )

        fig_scatter.update_layout(yaxis=dict(tickprefix="₹ ", tickformat=",.0f"))
        st.plotly_chart(fig_scatter, use_container_width=True)

        with st.expander("View Detailed Performance Table"):
            st.dataframe(filtered_df, use_container_width=True)

    except Exception as e:
        st.error(f"Error processing file: {e}")

else:
    st.info("Please upload a Sales Executive Excel file to generate the dashboard.")
