import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Sales Dashboard", layout="wide")

# -----------------------------
# INR Formatter
# -----------------------------
def format_inr(amount):
    try:
        return f"₹{float(amount):,.2f}"
    except:
        return "₹0.00"

# -----------------------------
# Detect & Clean File
# -----------------------------
@st.cache_data
def load_and_detect(file):
    df = pd.read_excel(file)

    df.columns = [c.strip().lower() for c in df.columns]

    has_sales_person = any("sales" in c for c in df.columns)

    if has_sales_person:
        mode = "sales"

        rename_map = {}
        for c in df.columns:
            if "sales" in c:
                rename_map[c] = "Sales Person Name"
            elif "night" in c:
                rename_map[c] = "Nights"
            elif "revenue" in c or "amount" in c:
                rename_map[c] = "Room Revenue"
            elif "pax" in c or "guest" in c:
                rename_map[c] = "Pax"
            elif "arr" in c:
                rename_map[c] = "ARR"

        df = df.rename(columns=rename_map)

    else:
        mode = "date"

        rename_map = {}
        for c in df.columns:
            if "date" in c:
                rename_map[c] = "Date"
            elif "night" in c:
                rename_map[c] = "Nights"
            elif "revenue" in c or "amount" in c:
                rename_map[c] = "Room Revenue"
            elif "pax" in c or "guest" in c:
                rename_map[c] = "Pax"

        df = df.rename(columns=rename_map)

    if "Room Revenue" not in df.columns:
        raise ValueError("Room Revenue column not found")

    for col in ["Nights", "Room Revenue", "Pax", "ARR"]:
        if col in df.columns:
            df[col] = (
                df[col].astype(str)
                .str.replace("₹", "")
                .str.replace(",", "")
            )
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df, mode

# -----------------------------
# UI
# -----------------------------
st.title("Sales Performance Dashboard")

uploaded_file = st.file_uploader("Upload Excel File", type=["xls", "xlsx"])

if uploaded_file:
    try:
        df, mode = load_and_detect(uploaded_file)

        st.sidebar.header("Filters")

        if mode == "sales":
            filter_col = "Sales Person Name"
            title_suffix = "Sales Executive Wise"
        else:
            filter_col = "Date"
            title_suffix = "Date Wise"

        selected = st.sidebar.multiselect(
            f"Select {filter_col}",
            df[filter_col].astype(str).unique(),
            default=df[filter_col].astype(str).unique()
        )

        filtered_df = df[df[filter_col].astype(str).isin(selected)]

        st.subheader(f"Performance Analysis ({title_suffix})")
        st.divider()

        total_revenue = filtered_df["Room Revenue"].sum()
        total_nights = int(filtered_df["Nights"].sum()) if "Nights" in filtered_df else 0
        total_pax = int(filtered_df["Pax"].sum()) if "Pax" in filtered_df else 0
        avg_arr = total_revenue / total_nights if total_nights > 0 else 0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Room Revenue", format_inr(total_revenue))
        c2.metric("Total Nights Sold", total_nights)
        c3.metric("Average Room Rate (ARR)", format_inr(avg_arr))
        c4.metric("Total Pax", total_pax)

        st.divider()

        fig_bar = px.bar(
            filtered_df,
            x=filter_col,
            y="Room Revenue",
            title=f"Revenue by {filter_col}"
        )
        fig_bar.update_layout(yaxis_tickprefix="₹ ")
        st.plotly_chart(fig_bar, use_container_width=True)

        if "Nights" in filtered_df.columns:
            fig_scatter = px.scatter(
                filtered_df,
                x="Nights",
                y="Room Revenue",
                color=filter_col,
                title="Nights vs Revenue"
            )
            st.plotly_chart(fig_scatter, use_container_width=True)

        with st.expander("View Data"):
            st.dataframe(filtered_df)

    except Exception as e:
        st.error(f"Error processing file: {e}")

else:
    st.info("Upload your Excel file to view dashboard.")
