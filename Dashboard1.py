import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Page Configuration
st.set_page_config(page_title="Sales Executive Performance Dashboard - Rhythm Gurugram", layout="wide")

# 2. Data Cleaning & Preparation
@st.cache_data
def load_cleaned_data():
    file_path = 'C:\\Users\\SC\\Desktop\\WINHMS Data Mining\\Data\\Sales Executive Dec.xls'
    
    # Skip the first 3 rows of metadata to reach the header row
    df = pd.read_excel(file_path, skiprows=3)
    
    # Remove the first empty column and rename columns
    # We rename 'Company' to 'Sales Person Name' as requested
    df = df.iloc[:, 1:] 
    df.columns = ['Sales Person Name', 'Nights', 'Occupancy %', 'Pax', 'Room Revenue', 'Revenue%', 'ARR', 'ARP']
    
    # Data Cleaning: Remove empty rows and exclude 'Total' / 'Not Defined' rows
    df = df.dropna(subset=['Sales Person Name'])
    df = df[~df['Sales Person Name'].str.contains('Total|Not Defined|Grand Total', case=False, na=False)]
    
    # Convert numeric columns to actual numbers for calculation
    numeric_cols = ['Nights', 'Occupancy %', 'Pax', 'Room Revenue', 'Revenue%', 'ARR', 'ARP']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
    return df

try:
    df = load_cleaned_data()

    # 3. Sidebar Filters
    st.sidebar.header("Dashboard Controls")
    selected_names = st.sidebar.multiselect(
        "Select Sales Persons", 
        options=df['Sales Person Name'].unique(), 
        default=df['Sales Person Name'].unique()
    )

    filtered_df = df[df['Sales Person Name'].isin(selected_names)]

    # 4. Header & KPIs
    st.title("Sales Executive Performance Dashboard - Rhythm Gurugram")
    st.subheader("Performance Analysis for December 2025")
    st.divider()

    # Formatted KPI Cards with Indian Rupee Style
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Total Room Revenue", f"₹{filtered_df['Room Revenue'].sum():,.2f}")
    kpi2.metric("Total Nights Sold", int(filtered_df['Nights'].sum()))
    kpi3.metric("Avg. Room Rate (ARR)", f"₹{filtered_df['ARR'].mean():,.2f}")
    kpi4.metric("Total Pax (Guests)", int(filtered_df['Pax'].sum()))

    st.divider()

    # 5. Visualizations
    col_left, col_right = st.columns(2)

    with col_left:
        st.subheader("Revenue by Sales Person")
        # Creating a horizontal bar chart
        fig_rev = px.bar(
            filtered_df.sort_values('Room Revenue', ascending=True),
            x='Room Revenue',
            y='Sales Person Name',
            orientation='h',
            color='Room Revenue',
            color_continuous_scale='Blues',
            # This ensures the hover label says "Sales Person Name"
            labels={'Room Revenue': 'Room Revenue (₹)', 'Sales Person Name': 'Sales Person Name'}
        )
        # Force the X-axis to show Rupees and full numbers
        fig_rev.update_layout(
            xaxis=dict(tickprefix="₹ ", tickformat=",.0f"),
            yaxis={'categoryorder':'total ascending'}
        )
        st.plotly_chart(fig_rev, use_container_width=True)

    with col_right:
        st.subheader("Revenue Contribution (%)")
        fig_pie = px.pie(
            filtered_df,
            values='Room Revenue',
            names='Sales Person Name',
            hole=0.4,
            # Hover labels set here as well
            labels={'Sales Person Name': 'Sales Person Name', 'Room Revenue': 'Revenue (₹)'}
        )
        fig_pie.update_traces(textinfo='percent+label')
        st.plotly_chart(fig_pie, use_container_width=True)

    # 6. Simplified Efficiency Chart
    st.divider()
    st.subheader("Efficiency Analysis: Nights vs. Revenue")
    st.write("This chart helps you see who is selling more nights versus who is bringing in higher value revenue.")
    
    fig_scatter = px.scatter(
        filtered_df,
        x='Nights',
        y='Room Revenue',
        size='Room Revenue',
        color='Sales Person Name',
        hover_name='Sales Person Name',
        labels={'Room Revenue': 'Room Revenue (₹)', 'Nights': 'Nights Sold'}
    )
    fig_scatter.update_layout(yaxis=dict(tickprefix="₹ ", tickformat=",.0f"))
    st.plotly_chart(fig_scatter, use_container_width=True)

    # 7. Cleaned Data Table for reference
    with st.expander("View Detailed Performance Table"):
        # Formatting the dataframe for display
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
    st.info("Please make sure the CSV file is in the same folder as this script.")