import streamlit as st
import pandas as pd
import plotly.express as px

# Set Page Config
st.set_page_config(page_title="Rhythm Gurugram Sales Dashboard", layout="wide")

# Load Data
@st.cache_data
def load_data():
    # Assuming the consolidated file exists from previous step
    df = pd.read_csv('consolidated_sales_data.csv')
    
    # Define chronological order for months
    month_order = [
        'April 25', 'May 25', 'June 25', 'July 25', 'Aug 25', 
        'Sep 25', 'Oct 25', 'Nov 25', 'Dec 25', 'Jan 26'
    ]
    df['Month-Year'] = pd.Categorical(df['Month-Year'], categories=month_order, ordered=True)
    return df

df = load_data()

# Sidebar - Title and Filters
st.sidebar.header("Dashboard Filters")

# Sales Executive Slicer
executives = ["All"] + sorted(df['Sales Executive'].unique().tolist())
selected_executive = st.sidebar.selectbox("Select Sales Executive", executives)

# Month Slicer
months = ["All"] + df['Month-Year'].cat.categories.tolist()
selected_month = st.sidebar.selectbox("Select Month", months)

# Filtering Logic
filtered_df = df.copy()
if selected_executive != "All":
    filtered_df = filtered_df[filtered_df['Sales Executive'] == selected_executive]
if selected_month != "All":
    filtered_df = filtered_df[filtered_df['Month-Year'] == selected_month]

# Dashboard Heading
st.title("Rhythm Gurugram Sales Productivity Dashboard")
st.markdown("---")

# KPI Metrics
total_room_nights = filtered_df['Room Nights'].sum()
total_revenue = filtered_df['Revenue'].sum()
avg_arr = filtered_df['ARR'].mean() if not filtered_df.empty else 0

col1, col2, col3 = st.columns(3)
with col1:
    st.metric(label="Total Room Nights", value=f"{total_room_nights:,.0f}")
with col2:
    st.metric(label="Total Revenue", value=f"₹{total_revenue:,.2f}")
with col3:
    st.metric(label="Average ARR", value=f"₹{avg_arr:,.2f}")

st.markdown("---")

# Visualization Section
row1_col1, row1_col2 = st.columns(2)

with row1_col1:
    st.subheader("Monthly Revenue Trend")
    monthly_rev = filtered_df.groupby('Month-Year', observed=True)['Revenue'].sum().reset_index()
    
    # Removed '.2s' to show full numbers and added comma separators
    fig_rev = px.bar(monthly_rev, x='Month-Year', y='Revenue', 
                     text_auto=',.2f', 
                     color_discrete_sequence=['#1f77b4'],
                     labels={'Revenue': 'Revenue (₹)', 'Month-Year': 'Month'})
    
    # Disable SI prefixes (k, M) on the Y-axis
    fig_rev.update_layout(yaxis=dict(tickformat=",.0f"), xaxis_title="Month", yaxis_title="Revenue (Full Value)")
    st.plotly_chart(fig_rev, use_container_width=True)

with row1_col2:
    st.subheader("Monthly Room Nights")
    monthly_rn = filtered_df.groupby('Month-Year', observed=True)['Room Nights'].sum().reset_index()
    
    fig_rn = px.line(monthly_rn, x='Month-Year', y='Room Nights', markers=True,
                     text='Room Nights',
                     labels={'Room Nights': 'Nights', 'Month-Year': 'Month'})
    
    # Formatting line chart labels and axis
    fig_rn.update_traces(textposition="top center")
    fig_rn.update_layout(yaxis=dict(tickformat=",.0f"), xaxis_title="Month", yaxis_title="Room Nights")
    st.plotly_chart(fig_rn, use_container_width=True)

row2_col1, row2_col2 = st.columns(2)

with row2_col1:
    st.subheader("Revenue Contribution by Executive")
    if selected_executive == "All":
        exec_rev = filtered_df.groupby('Sales Executive')['Revenue'].sum().reset_index()
        # Showing full values in the pie chart labels
        fig_pie = px.pie(exec_rev, values='Revenue', names='Sales Executive', hole=0.4)
        fig_pie.update_traces(textinfo='percent+value', texttemplate='%{label}: <br>₹%{value:,.2f}')
        st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.info(f"Detailed Monthly Performance for: {selected_executive}")
        st.dataframe(filtered_df[['Month-Year', 'Room Nights', 'Revenue', 'ARR']].sort_values('Month-Year'))

with row2_col2:
    st.subheader("ARR Performance by Month")
    fig_arr = px.scatter(filtered_df, x='Month-Year', y='ARR', size='Revenue', color='Sales Executive',
                         hover_name='Sales Executive')
    
    # Ensure ARR scale shows full numbers
    fig_arr.update_layout(yaxis=dict(tickformat=",.0f"), yaxis_title="ARR (₹)")
    st.plotly_chart(fig_arr, use_container_width=True)

# Data Table
with st.expander("View Consolidated Data Table"):
    st.write(filtered_df)