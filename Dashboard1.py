import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Sales Executive Performance Dashboard", layout="wide")

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
    return f"₹{integer}.{parts[1]}"

# -----------------------------
# Smart cleaner with fallback
# -----------------------------
def smart_clean_data(df):

    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]
    df.columns = df.columns.astype(str).str.strip().str.lower()

    mapped = {}

    for col in df.columns:
        c = col.lower()
        if "sales" in c or "executive" in c or "name" in c:
            mapped[col] = "Sales Person Name"
        elif "night" in c:
            mapped[col] = "Nights"
        elif "occupancy" in c:
            mapped[col] = "Occupancy %"
        elif "pax" in c or "guest" in c:
            mapped[col] = "Pax"
        elif any(x in c for x in ["revenue", "amount", "total", "value", "net"]):
            mapped[col] = "Room Revenue"
        elif "arr" in c:
            mapped[col] = "ARR"
        elif "arp" in c:
            mapped[col] = "ARP"
        elif "%" in c:
            mapped[col] = "Revenue%"

    df = df.rename(columns=mapped)

    return df

def finalize_dataframe(df):

    if "Sales Person Name" not in df.columns:
        raise ValueError("Sales Person Name column not found")

    if "Room Revenue" not in df.columns:
        return None  # trigger manual selection

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

    num_cols = ["Nights", "Occupancy %", "Pax", "Room Revenue", "Revenue%", "ARR", "ARP"]
    for col in num_cols:
        df[col] = (
            df[col].astype(str)
            .str.replace("₹", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.strip()
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df = df[~df["Sales Person Name"].astype(str).str.contains("total|grand", case=False, na=False)]

    return df

# -----------------------------
# KPI Styling
# -----------------------------
st.markdown("""
<style>
.kpi-box{padding:20px;border-radius:12px;background:#f6f8fa;text-align:center;box-shadow:0 2px 6px rgba(0,0,0,.05)}
.kpi-title{font-size:15px;color:#666}
.kpi-value{font-size:24px;font-weight:700;white-space:nowrap}
</style>
""", unsafe_allow_html=True)

st.title("Sales Executive Performance Dashboard")

uploaded_file = st.file_uploader("Upload Excel file", type=["xls", "xlsx"])

if uploaded_file:
    try:
        raw_df = pd.read_excel(uploaded_file)
        df = smart_clean_data(raw_df)

        df_final = finalize_dataframe(df)

        # -------- MANUAL REVENUE SELECTION --------
        if df_final is None:
            st.warning("Room Revenue column not detected automatically.")

            revenue_col = st.selectbox(
                "Select the column that represents Room Revenue:",
                raw_df.columns
            )

            df["Room Revenue"] = raw_df[revenue_col]
            df_final = finalize_dataframe(df)

        df = df_final

        st.sidebar.header("Filters")
        selected = st.sidebar.multiselect(
            "Sales Persons",
            df["Sales Person Name"].unique(),
            default=df["Sales Person Name"].unique()
        )

        fdf = df[df["Sales Person Name"].isin(selected)]

        total_revenue = fdf["Room Revenue"].sum()
        total_nights = int(fdf["Nights"].sum())
        avg_arr = fdf["ARR"].mean()
        total_pax = int(fdf["Pax"].sum())

        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f"<div class='kpi-box'><div class='kpi-title'>Total Revenue</div><div class='kpi-value'>{format_inr(total_revenue)}</div></div>", unsafe_allow_html=True)
        c2.markdown(f"<div class='kpi-box'><div class='kpi-title'>Total Nights</div><div class='kpi-value'>{total_nights}</div></div>", unsafe_allow_html=True)
        c3.markdown(f"<div class='kpi-box'><div class='kpi-title'>Average ARR</div><div class='kpi-value'>{format_inr(avg_arr)}</div></div>", unsafe_allow_html=True)
        c4.markdown(f"<div class='kpi-box'><div class='kpi-title'>Total Pax</div><div class='kpi-value'>{total_pax}</div></div>", unsafe_allow_html=True)

        st.divider()

        col1, col2 = st.columns(2)

        with col1:
            fig = px.bar(fdf.sort_values("Room Revenue"), x="Room Revenue", y="Sales Person Name", orientation="h")
            fig.update_layout(xaxis_tickprefix="₹ ")
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            fig2 = px.pie(fdf, values="Room Revenue", names="Sales Person Name", hole=0.4)
            st.plotly_chart(fig2, use_container_width=True)

        fig3 = px.scatter(fdf, x="Nights", y="Room Revenue", size="Room Revenue", color="Sales Person Name")
        fig3.update_layout(yaxis_tickprefix="₹ ")
        st.plotly_chart(fig3, use_container_width=True)

        with st.expander("View Data"):
            st.dataframe(fdf, use_container_width=True)

    except Exception as e:
        st.error(f"Error processing file: {e}")

else:
    st.info("Upload an Excel file to begin.")
