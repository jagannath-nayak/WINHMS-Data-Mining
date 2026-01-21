import streamlit as st
import pandas as pd
import plotly.express as px

# -----------------------------
# Page Configuration
# -----------------------------
st.set_page_config(
    page_title="Sales Executive Performance Dashboard - Rhythm Gurugram",
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
# Load & Clean Data
# -----------------------------
@st.cache_data
def load_cleaned_data():
    file_path = "data/Sales Executive Dec.xls"

    df = pd.read_excel(file_path, skiprows=3)
    df = df.iloc[:, 1:]

    df.columns = [
        'Sales Person Name', 'Nights', 'Occupancy %', 'Pax',
        'Room Revenue', 'Revenue%', 'ARR', 'ARP'
    ]

    df = df.dropna(subset=['Sales Person Name'])
    df = df[~df['Sales Person Name'].str.contains(
        'Total|Not Defined|Grand Total', case=False, na=False
    )]

    numeric_cols = ['Nights', 'Occupancy %', 'Pax', 'Room Revenue', 'Revenue%', 'ARR', 'ARP']

    for col in numeric_cols:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace("₹", "", regex=False)
            .str.replace(",", "", regex=False)
            .str.strip()
        )
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df

# -----------------------------
# KPI Card Styling (FIXED)
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
    white-space: nowrap;
}
</style>
""", unsafe_allow_html=True)

# -----------------------------
# Main App
# -----------------------------
try:
    df = load_cleaned_data()

    # Sidebar
    st.sidebar.header("Dashboard Controls")
    selected_names = st.sidebar.multiselect(
        "Select Sales Persons",
        options=df['Sales Person Name'].unique(),
        default=df['Sales Person Name'].unique()
    )

    filtered_df = df[df['Sales Person Name'].isin(selected_names)]

    # Header
    st.title("Sales Executive Performance Dashboard - Rhythm Gurugram")
    st.subheader("Performance Analysis for December 2025")
    st.divider()

    # KPI Calculations
    total_revenue = filtered_df['Room Revenue'].sum()
    total_nights = int(filtered_df['Nights'].sum())
    avg_arr = filtered_df['ARR'].mean()
    total_pax = int(filtered_df['Pax'].sum())

    # KPI Cards
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown(f"""
        <div class="kpi-box">
            <div class="kpi-title">Total Room Revenue</div>
            <div class="kpi-value">{format_inr(total_revenue)}</div>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown(f"""
        <div class="kpi-box">
            <div class="kpi-title">Total Nights Sold</div>
            <div class="kpi-value">{total_nights}</div>
        </div>
        """, unsafe_allow_html=True)

    with c3:
        st.markdown(f"""
        <div class="kpi-box">
            <div class="kpi-title">Average Room Rate (ARR)</div>
            <div class="kpi-value">{format_inr(avg_arr)}</div>
        </div>
        """, unsafe_allow_html=True)

    with c4:
        st.markdown(f"""
        <div class="kpi-box">
            <div class="kpi-title">Total Pax (Guests)</div>
            <div class="kpi-value">{total_pax}</div>
        </div>
        """, unsafe_allow_html=True)

    st.divider()

    # Charts
    col_left, col_right = st.columns(2)

    with col_left:
        st.subheader("Revenue by Sales Person")

        fig_rev = px.bar(
            filtered_df.sort_values('Room Revenue', ascending=True),
            x='Room Revenue',
            y='Sales Person Name',
            orientation='h',
            color='Room Revenue',
            color_continuous_scale='Blues'
        )

        fig_rev.update_layout(
            xaxis=dict(tickprefix="₹ ", tickformat=",.0f"),
            yaxis={'categoryorder': 'total ascending'}
        )

        st.plotly_chart(fig_rev, use_container_width=True)

    with col_right:
        st.subheader("Revenue Contribution (%)")

        fig_pie = px.pie(
            filtered_df,
            values='Room Revenue',
            names='Sales Person Name',
            hole=0.4
        )

        fig_pie.update_traces(textinfo='percent+label')
        st.plotly_chart(fig_pie, use_container_width=True)

    # Scatter
    st.divider()
    st.subheader("Efficiency Analysis: Nights vs Revenue")

    fig_scatter = px.scatter(
        filtered_df,
        x='Nights',
        y='Room Revenue',
        size='Room Revenue',
        color='Sales Person Name',
        hover_name='Sales Person Name'
    )

    fig_scatter.update_layout(yaxis=dict(tickprefix="₹ ", tickformat=",.0f"))
    st.plotly_chart(fig_scatter, use_container_width=True)

    # Data Table
    with st.expander("View Detailed Performance Table"):
        st.dataframe(
            filtered_df.style.format({
                'Room Revenue': '₹{:,.2f}',
                'ARR': '₹{:,.2f}',
                'ARP': '₹{:,.2f}',
                'Occupancy %': '{:.2f}%',
                'Revenue%': '{:.2f}%'
            }),
            use_container_width=True
        )

except Exception as e:
    st.error(f"Error: {e}")
    st.info("Please make sure the Excel file exists at: data/Sales Executive Dec.xls")
