import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Page Configuration
st.set_page_config(page_title="LNL Sales Executive Dashboard", layout="wide")

# 2. Data Cleaning & Preparation
@st.cache_data
def load_lnl_data():
    file_path = 'C:\\Users\\SC\\Desktop\\WINHMS Data Mining\\Data\\Sales Executive Wise LNL.xls'
    
    # Skip metadata rows (1-3) to get to headers
    df = pd.read_excel(file_path, skiprows=3)
    
    # Remove first empty column and rename columns
    df = df.iloc[:, 1:] 
    df.columns = ['Sales Person Name', 'Nights', 'Occupancy %', 'Pax', 'Room Revenue', 'Revenue%', 'ARR', 'ARP']
    
    # Cleaning: Remove null names and exclude 'Total' rows
    df = df.dropna(subset=['Sales Person Name'])
    df = df[~df['Sales Person Name'].str.contains('Total|Grand Total|Not Defined', case=False, na=False)]
    
    # Convert numeric columns to float
    numeric_cols = ['Nights', 'Occupancy %', 'Pax', 'Room Revenue', 'Revenue%', 'ARR', 'ARP']
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
    return df

try:
    df = load_lnl_data()

    # 3. Sidebar Filter
    st.sidebar.header("Dashboard Filters")
    all_names = sorted(df['Sales Person Name'].unique())
    selected_names = st.sidebar.multiselect(
        "Filter by Sales Person:", 
        options=all_names, 
        default=all_names
    )

    filtered_df = df[df['Sales Person Name'].isin(selected_names)]

    # 4. Header
    st.title("Sales Executive Performance Dashboard - Rhythm Lonavala")
    st.markdown("### Performance Overview | December 2025")
    st.divider()

    # 5. FIXED TOTAL REVENUE (Removing "...")
    # Instead of standard st.metric which might truncate, we use markdown for a big, clear display
    total_rev = filtered_df['Room Revenue'].sum()
    total_nights = int(filtered_df['Nights'].sum())
    avg_arr = filtered_df['ARR'].mean()

    col1, col2, col3 = st.columns([2, 1, 1])
    
    with col1:
        st.markdown(f"**Total Room Revenue**")
        # Displaying the full number with ₹ and commas to avoid "..."
        st.markdown(f"<h1 style='color: #1E88E5;'>₹ {total_rev:,.2f}</h1>", unsafe_allow_html=True)
    
    with col2:
        st.metric("Total Nights Sold", f"{total_nights}")
    
    with col3:
        st.metric("Average ARR", f"₹{avg_arr:,.2f}")

    st.divider()

    # 6. BIGGER & CLEARER CONTRIBUTION CHART
    # We use a Treemap here because it is much easier to read than a Pie chart for many people.
    st.subheader("📊 Revenue Contribution Analysis")
    st.write("The size of the box represents the share of total revenue. Click on a box to zoom in.")
    
    fig_tree = px.treemap(
        filtered_df, 
        path=['Sales Person Name'], 
        values='Room Revenue',
        color='Room Revenue',
        color_continuous_scale='Blues',
        labels={'Room Revenue': 'Revenue (₹)', 'Sales Person Name': 'Sales Person Name'}
    )
    
    # Customizing the hover and labels to show Indian Rupees
    fig_tree.update_traces(
        textinfo="label+value+percent root",
        texttemplate="<b>%{label}</b><br>₹%{value:,.2f}<br>%{percentRoot:.1%}",
        hovertemplate="<b>%{label}</b><br>Revenue: ₹%{value:,.2f}<br>Share: %{percentRoot:.1%}"
    )
    
    # Making the chart height bigger
    fig_tree.update_layout(height=500, margin=dict(t=10, l=10, r=10, b=10))
    st.plotly_chart(fig_tree, use_container_width=True)

    st.divider()

    # 7. PERFORMANCE DETAILS
    left, right = st.columns(2)

    with left:
        st.subheader("Revenue by Person (List)")
        fig_bar = px.bar(
            filtered_df.sort_values('Room Revenue', ascending=True),
            x='Room Revenue',
            y='Sales Person Name',
            orientation='h',
            text_auto=',.0f',
            labels={'Room Revenue': 'Revenue (₹)', 'Sales Person Name': 'Sales Person Name'}
        )
        fig_bar.update_layout(xaxis=dict(tickprefix="₹ ", tickformat=",.0f"))
        st.plotly_chart(fig_bar, use_container_width=True)

    with right:
        st.subheader("Nights vs. Revenue")
        fig_scatter = px.scatter(
            filtered_df,
            x='Nights',
            y='Room Revenue',
            size='Room Revenue',
            color='Sales Person Name',
            hover_name='Sales Person Name',
            labels={'Room Revenue': 'Revenue (₹)', 'Nights': 'Nights Sold'}
        )
        fig_scatter.update_layout(yaxis=dict(tickprefix="₹ ", tickformat=",.0f"))
        st.plotly_chart(fig_scatter, use_container_width=True)

    # 8. DATA TABLE
    with st.expander("View Full Data Sheet"):
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