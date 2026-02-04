import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np

# Set page configuration
st.set_page_config(page_title="Rhythm Sales Dashboard", layout="wide")

# Custom CSS for Professional Layout
st.markdown("""
    <style>
    h1 { font-size: 26px !important; margin-bottom: 5px; color: #1f77b4; font-weight: bold; }
    .kpi-container {
        background-color: #1E1E2F;
        color: white;
        padding: 20px;
        border-radius: 12px;
        margin-bottom: 25px;
    }
    .kpi-label { font-size: 13px; color: #ADB5BD; text-transform: uppercase; }
    .kpi-value { font-size: 26px; font-weight: bold; color: #00D1FF; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏨 Rhythm Hotels Sales Productivity Dashboard")

def process_hotel_data(uploaded_file):
    try:
        # Load the file
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            df_raw = pd.read_excel(uploaded_file, header=None)
        
        # 1. Find the exact row indices for Header and Total
        header_row_idx = None
        total_row_idx = None
        
        for i, row in df_raw.iterrows():
            row_vals = [str(val).strip().upper() for val in row.values]
            if 'COMPANY' in row_vals:
                header_row_idx = i
            if 'TOTAL' in row_vals and i > (header_row_idx if header_row_idx is not None else 0):
                total_row_idx = i
                break
        
        if header_row_idx is None or total_row_idx is None:
            st.error("Error: Could not find 'Company' header or 'Total' summary row.")
            return None, None

        # Helper to safely clean numeric strings
        def safe_float(val):
            try:
                # Remove commas and non-numeric characters except decimals
                clean_val = "".join(c for c in str(val) if c.isdigit() or c in '.-')
                return float(clean_val) if clean_val else 0.0
            except:
                return 0.0

        # 2. Extract Master Totals from the Summary row
        # (Assuming Col index 1 for Nights and index 4 or 5 for Revenue based on your files)
        total_row = df_raw.iloc[total_row_idx]
        
        # We try to find the numbers in the total row automatically
        numeric_values_in_total = [safe_float(x) for x in total_row.values if safe_float(x) > 0]
        
        # Nights is usually the smaller number, Revenue is the larger
        m_nights = min(numeric_values_in_total) if numeric_values_in_total else 0
        m_revenue = max(numeric_values_in_total) if numeric_values_in_total else 0
        
        # Correct ARR Formula
        m_arr = m_revenue / m_nights if m_nights > 0 else 0
        
        master_kpis = {
            "revenue": int(round(m_revenue)),
            "nights": int(round(m_nights)),
            "arr": int(round(m_arr))
        }

        # 3. Extract Company data (Slicing strictly BETWEEN header and total)
        df = df_raw.iloc[header_row_idx + 1 : total_row_idx].copy()
        df.columns = [str(val).strip() for val in df_raw.iloc[header_row_idx]]
        
        # Map required columns
        col_map = {'Company': 'Company', 'Nights': 'Nights', 'Room Revenue': 'Room_Revenue'}
        df = df[[c for c in df.columns if c in col_map]].rename(columns=col_map)
        
        # Clean Data Types
        df['Company'] = df['Company'].astype(str).str.strip()
        for col in ['Nights', 'Room_Revenue']:
            df[col] = df[col].apply(safe_float)
                
        # Calculate ARR per company
        df['ARR'] = np.where(df['Nights'] > 0, df['Room_Revenue'] / df['Nights'], 0)
        
        # Convert to int for display (removes the "int base 10" error risk)
        for col in ['Nights', 'Room_Revenue', 'ARR']:
            df[col] = df[col].apply(lambda x: int(round(x)))
            
        return df, master_kpis

    except Exception as e:
        st.error(f"Processing Error: {str(e)}")
        return None, None

# --- SIDEBAR & DISPLAY ---
uploaded_file = st.sidebar.file_uploader("Upload File", type=["xls", "xlsx", "csv"])

if uploaded_file:
    df, master_kpis = process_hotel_data(uploaded_file)
    
    if df is not None:
        st.markdown(f"""
            <div class="kpi-container">
                <div style="display: flex; justify-content: space-around; text-align: center;">
                    <div><div class="kpi-label">Total Room Revenue</div><div class="kpi-value">₹ {master_kpis['revenue']:,}</div></div>
                    <div><div class="kpi-label">Total Room Nights</div><div class="kpi-value">{master_kpis['nights']:,}</div></div>
                    <div><div class="kpi-label">Average Room Rate (ARR)</div><div class="kpi-value">₹ {master_kpis['arr']:,}</div></div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        all_companies = sorted(df['Company'].unique())
        selected = st.sidebar.multiselect("Select Companies:", options=all_companies, default=df.nlargest(10, 'Nights')['Company'].tolist())
        
        if selected:
            filtered_df = df[df['Company'].isin(selected)]
            st.plotly_chart(px.bar(filtered_df.sort_values('Nights'), x='Nights', y='Company', orientation='h', title="Nights Comparison"), use_container_width=True)
            st.subheader("📋 Performance Table")
            st.dataframe(filtered_df.sort_values('Room_Revenue', ascending=False), use_container_width=True, hide_index=True)