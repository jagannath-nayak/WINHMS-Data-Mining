import streamlit as st
import pandas as pd
import plotly.express as px

# -----------------------------
# Page Config
# -----------------------------
st.set_page_config(page_title="WINHMS Sales Executive Dashboard", layout="wide")

st.title("🏨 WINHMS Sales Executive Dashboard")
st.write("Upload Sales Executive Wise Excel report")

uploaded_file = st.file_uploader("Upload Excel file", type=["xls", "xlsx"])

# -----------------------------
# Data Loader
# -----------------------------
@st.cache_data
def load_data(file):

    df = pd.read_excel(file, skiprows=3)

    # Drop empty columns
    df = df.dropna(axis=1, how="all")

    # Drop first unnamed column if exists
    if "Unnamed: 0" in df.columns:
        df = df.drop(columns=["Unnamed: 0"])

    # Validate column count
    if len(df.columns) < 8:
        raise ValueError("Invalid WINHMS file. Please upload Sales Executive Wise LNL report.")

    df = df.iloc[:, :8]
    df.columns = [
        'Sales Person Name', 'Nights', 'Occupancy %', 'Pax',
        'Room Revenue', 'Revenue%', 'ARR', 'ARP'
    ]

    # Convert name to string
    df['Sales Person Name'] = df['Sales Person Name'].astype(str)

    # Remove total rows
    df = df[~df['Sales Person Name'].str.contains(
        'total|grand|not defined', case=False, na=False
    )]

    # Convert numeric columns
    num_cols = ['Nights', 'Occupancy %', 'Pax', 'Room Revenue', 'Revenue%', 'ARR', 'ARP']
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Remove NaN rows
    df = df.dropna(subset=['Sales Person Name', 'Room Revenue'])

    # Remove zero revenue rows
    df = df[df['Room Revenue'] > 0]

    # Reset index
    df = df.reset_index(drop=True)

    return df


# -----------------------------
# Dashboard
# -----------------------------
if uploaded_file:

    try:
        df = load_data(uploaded_file)

        # Sidebar Filter
        st.sidebar.header("Filters")
        people = sorted(df['Sales Person Name'].unique())

        selected = st.sidebar.multiselect("Sales Executive", people, default=people)
        filtered_df = df[df['Sales Person Name'].isin(selected)]

        # KPIs
        total_revenue = filtered_df['Room Revenue'].sum()
        total_nights = int(filtered_df['Nights'].sum())
        avg_arr = filtered_df['ARR'].mean()

        col1, col2, col3 = st.columns(3)
        col1.metric("Total Revenue", f"₹ {total_revenue:,.2f}")
        col2.metric("Total Nights", total_nights)
        col3.metric("Average ARR", f"₹ {avg_arr:,.2f}")

        st.divider()

        # -----------------------------
        # Revenue Contribution Treemap
        # -----------------------------
        fig_tree = px.treemap(
            filtered_df,
            path=['Sales Person Name'],
            values='Room Revenue',
            title="Revenue Contribution (₹)"
        )

        fig_tree.update_traces(
            texttemplate="<b>%{label}</b><br>₹%{value:,.0f}<br>%{percentRoot:.1%}",
            hovertemplate="<b>%{label}</b><br>Revenue: ₹%{value:,.2f}<br>Share: %{percentRoot:.1%}"
        )

        st.plotly_chart(fig_tree, use_container_width=True)

        # -----------------------------
        # Revenue Bar Chart
        # -----------------------------
        fig_bar = px.bar(
            filtered_df.sort_values("Room Revenue"),
            x="Room Revenue",
            y="Sales Person Name",
            orientation="h",
            title="Revenue by Sales Executive (₹)"
        )

        fig_bar.update_layout(
            xaxis=dict(tickprefix="₹ ", tickformat=",.0f")
        )

        st.plotly_chart(fig_bar, use_container_width=True)

        # -----------------------------
        # Scatter Chart
        # -----------------------------
        fig_scatter = px.scatter(
            filtered_df,
            x="Nights",
            y="Room Revenue",
            size="Room Revenue",
            color="Sales Person Name",
            title="Nights vs Revenue (₹)"
        )

        fig_scatter.update_layout(
            yaxis=dict(tickprefix="₹ ", tickformat=",.0f")
        )

        st.plotly_chart(fig_scatter, use_container_width=True)

        # -----------------------------
        # Data Table (clean)
        # -----------------------------
        with st.expander("📄 View Clean Data Table"):
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
        st.error("❌ Failed to process file")
        st.exception(e)

else:
    st.info("Please upload a WINHMS Sales Executive Wise LNL Excel file.")
