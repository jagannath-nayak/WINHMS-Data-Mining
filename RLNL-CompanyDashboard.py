import streamlit as st
import pandas as pd
import plotly.express as px
import numpy as np

# Set page configuration
st.set_page_config(page_title="Rhythm Lonavala Sales Dashboard", layout="wide")

# Custom CSS for Professional Layout
st.markdown("""
    <style>
    h1 { font-size: 26px !important; margin-bottom: 5px; color: #1f77b4; font-weight: bold; }
    h2 { font-size: 20px !important; margin-top: 10px; }
    
    .kpi-container {
        background-color: #1E1E2F;
        color: white;
        padding: 20px;
        border-radius: 12px;
        margin-bottom: 25px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .kpi-label { font-size: 13px; color: #ADB5BD; text-transform: uppercase; letter-spacing: 1px; }
    .kpi-value { font-size: 26px; font-weight: bold; color: #00D1FF; margin-top: 5px; }
    
    .main { background-color: #F8F9FA; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏨 Rhythm Lonavala Sales Productivity Dashboard")

def process_hotel_data(uploaded_file):
    try:
        # Load the file
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            df_raw = pd.read_excel(uploaded_file, header=None)
        
        # 1. Find Header and Total rows
        header_row_idx = None
        total_row_idx = None
        
        for i, row in df_raw.iterrows():
            row_vals = [str(val).strip().upper() for val in row.values]
            if 'COMPANY' in row_vals:
                header_row_idx = i
            if 'TOTAL' in row_vals:
                total_row_idx = i
                break
        
        if header_row_idx is None or total_row_idx is None:
            st.error("Error: Could not find 'Company' header or 'Total' row.")
            return None, None

        # 2. Extract Master Totals (Directly from source)
        total_row = df_raw.iloc[total_row_idx]
        m_nights = float(str(total_row[2]).replace(',', '').strip())
        m_revenue = float(str(total_row[5]).replace(',', '').strip())
        m_arr = m_revenue / m_nights if m_nights > 0 else 0
        
        master_kpis = {
            "revenue": int(round(m_revenue)),
            "nights": int(round(m_nights)),
            "arr": int(round(m_arr))
        }

        # 3. Extract Company data
        df = df_raw.iloc[header_row_idx:total_row_idx].copy()
        df.columns = [str(val).strip() for val in df.iloc[0]]
        df = df.iloc[1:].reset_index(drop=True)
        
        # Standardize and Clean
        col_map = {'Company': 'Company', 'Nights': 'Nights', 'Room Revenue': 'Room_Revenue'}
        df = df[[c for c in df.columns if c in col_map]].rename(columns=col_map)
        
        # FIX: Ensure Company is always a string to avoid TypeError in slicer
        df['Company'] = df['Company'].astype(str).str.strip()
        df = df[df['Company'] != 'nan']
        
        # Clean Numeric Columns
        for col in ['Nights', 'Room_Revenue']:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                
        # ARR = Revenue / Nights
        df['ARR'] = np.where(df['Nights'] > 0, df['Room_Revenue'] / df['Nights'], 0)
        
        # Rounding for presentation
        for col in ['Nights', 'Room_Revenue', 'ARR']:
            df[col] = df[col].round(0).astype(int)
            
        return df, master_kpis

    except Exception as e:
        st.error(f"Failed to process file: {str(e)}")
        return None, None

# --- FILE UPLOAD ---
st.sidebar.header("📁 Upload File")
uploaded_file = st.sidebar.file_uploader("Upload XLS or CSV", type=["xls", "xlsx", "csv"])

if uploaded_file:
    df, master_kpis = process_hotel_data(uploaded_file)
    
    if df is not None:
        # --- TOP KPIs ---
        st.markdown(f"""
            <div class="kpi-container">
                <div style="display: flex; justify-content: space-around; text-align: center;">
                    <div><div class="kpi-label">Total Room Revenue</div><div class="kpi-value">₹ {master_kpis['revenue']:,}</div></div>
                    <div style="border-left: 1px solid #464855; height: 50px;"></div>
                    <div><div class="kpi-label">Total Room Nights</div><div class="kpi-value">{master_kpis['nights']:,}</div></div>
                    <div style="border-left: 1px solid #464855; height: 50px;"></div>
                    <div><div class="kpi-label">Average Room Rate (ARR)</div><div class="kpi-value">₹ {master_kpis['arr']:,}</div></div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        # --- DYNAMIC SLICER ---
        st.sidebar.header("🔍 Slicer / Filters")
        # FIX: Explicitly cast to string and filter out any remaining non-string artifacts
        all_companies = sorted([str(c) for c in df['Company'].unique() if str(c) != 'nan'])
        
        top_10_nights = df.nlargest(10, 'Nights')['Company'].tolist()
        
        selected_companies = st.sidebar.multiselect(
            "Select Company Names:", 
            options=all_companies, 
            default=top_10_nights
        )
        
        if not selected_companies:
            st.warning("Please select at least one company from the sidebar.")
        else:
            filtered_df = df[df['Company'].isin(selected_companies)]

            col1, col2 = st.columns(2)
            with col1:
                fig_nts = px.bar(filtered_df.sort_values('Nights'), x='Nights', y='Company', orientation='h', 
                                 title="Room Nights (Selection)", color_discrete_sequence=['#1f77b4'])
                st.plotly_chart(fig_nts, use_container_width=True)

            with col2:
                fig_pie = px.pie(filtered_df, values='Room_Revenue', names='Company', 
                                 title="Revenue Share (%)", hole=0.45, color_discrete_sequence=px.colors.qualitative.G10)
                st.plotly_chart(fig_pie, use_container_width=True)

            st.subheader("📋 Performance Data Table")
            st.dataframe(filtered_df.sort_values('Room_Revenue', ascending=False), use_container_width=True, hide_index=True)
else:
    st.info("👋 Upload your analysis file to view the Dashboard.")