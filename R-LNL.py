import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ─── Page Config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Rhythm Lonavala | Sales Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Sky-Blue Theme CSS ───────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    .stApp { background: linear-gradient(135deg, #e8f4fd 0%, #dbeeff 50%, #e0f2fe 100%); }
    .main .block-container { background: transparent; padding-top: 1rem; }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #ffffff 0%, #f0f8ff 100%);
        border-right: 2px solid #b3d9f5;
    }
    [data-testid="stSidebar"] * { color: #1a3a5c !important; }

    .kpi-card {
        background: linear-gradient(135deg, #ffffff 0%, #f0f8ff 100%);
        border: 1.5px solid #90c8f0;
        border-radius: 14px;
        padding: 18px 16px;
        text-align: center;
        box-shadow: 0 4px 16px rgba(59, 130, 180, 0.15);
        margin-bottom: 4px;
    }
    .kpi-icon  { font-size: 1.6rem; margin-bottom: 4px; }
    .kpi-label {
        color: #4a7fa5; font-size: 0.72rem; font-weight: 600;
        text-transform: uppercase; letter-spacing: 1px; margin-bottom: 6px;
    }
    .kpi-value { color: #0d3b66; font-size: 1.4rem; font-weight: 700; line-height: 1.2; word-break: break-all; }
    .kpi-sub   { font-size: 0.72rem; color: #5b9bd5; margin-top: 4px; }

    .section-header {
        background: linear-gradient(90deg, #1e6fa8 0%, #3a9bd5 100%);
        color: #ffffff !important; font-size: 0.95rem; font-weight: 600;
        padding: 8px 16px; border-radius: 8px; margin: 20px 0 12px 0;
    }
    .rank-row {
        display: flex; align-items: center; padding: 10px 14px;
        border-radius: 10px; margin-bottom: 6px;
        background: #ffffff; border: 1.5px solid #b3d9f5;
        box-shadow: 0 2px 6px rgba(59,130,180,0.08);
    }
    .page-title { color: #0d3b66; font-size: 1.7rem; font-weight: 800; margin: 0 0 4px 0; }
    .page-sub   { color: #4a7fa5; font-size: 0.88rem; margin: 0 0 18px 0; }

    #MainMenu, footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ─── Constants ────────────────────────────────────────────────────────────────
PLOT_BG    = "#f8fbff"
PAPER_BG   = "#f8fbff"
FONT_COLOR = "#1a3a5c"
GRID_COLOR = "#d0e8f5"
COLORS = [
    "#1e6fa8","#e84393","#27ae60","#f39c12",
    "#8e44ad","#e74c3c","#16a085","#2980b9",
    "#d35400","#7f8c8d","#c0392b","#1abc9c",
    "#f1c40f","#9b59b6"
]

def fmt_inr(val):
    """Full Indian Rupee format — no K/L/Cr shortcuts."""
    if pd.isna(val) or val == 0:
        return "₹0"
    return f"₹{val:,.0f}"

# ─── Data Loading ─────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    df = pd.read_csv("consolidated_lonavala_sales_data.csv")
    df.columns = df.columns.str.strip()
    for col in ["Room Nights", "Revenue", "ARR"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    month_order = {
        "April 25": 1, "May 25": 2, "June 25": 3, "July 25": 4,
        "Aug 25": 5, "Sep 25": 6, "Oct 25": 7, "Nov 25": 8,
        "Dec 25": 9, "Jan 26": 10
    }
    df["Month_Order"] = df["Month-Year"].map(month_order)
    df = df.sort_values("Month_Order").reset_index(drop=True)
    return df

df = load_data()
all_months = df.sort_values("Month_Order")["Month-Year"].unique().tolist()
all_execs  = sorted(df["Sales Executive"].unique().tolist())

# ─── Sidebar — Dropdown Filters ───────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='text-align:center;padding:16px 0 10px 0;'>
        <div style='font-size:2.4rem;'></div>
        <div style='color:#0d3b66;font-size:1.2rem;font-weight:700;'>Rhythm Lonavala</div>
        <div style='color:#4a7fa5;font-size:0.72rem;letter-spacing:2px;text-transform:uppercase;'>Sales Dashboard</div>
    </div>
    <hr style='border-color:#b3d9f5;margin:10px 0;'>
    <div style='color:#4a7fa5;font-size:0.8rem;font-weight:600;margin-bottom:14px;'>🎛️ FILTERS</div>
    """, unsafe_allow_html=True)

    sel_month = st.selectbox(
        "📅 Select Month",
        options=["All Months"] + all_months,
        index=0
    )

    sel_exec = st.selectbox(
        "👤 Select Executive",
        options=["All Executives"] + all_execs,
        index=0
    )

    st.markdown("<hr style='border-color:#b3d9f5;margin:16px 0;'>", unsafe_allow_html=True)
    st.markdown("""
    <div style='color:#4a7fa5;font-size:0.72rem;text-align:center;line-height:1.9;'>
        📊 Data Period<br>
        <span style='color:#0d3b66;font-weight:700;'>April 2025 – January 2026</span><br><br>
    </div>
    """, unsafe_allow_html=True)

# ─── Apply Filters ────────────────────────────────────────────────────────────
mask = pd.Series([True] * len(df), index=df.index)
if sel_month != "All Months":
    mask &= df["Month-Year"] == sel_month
if sel_exec != "All Executives":
    mask &= df["Sales Executive"] == sel_exec

filtered   = df[mask].copy()
# Sales reps = exclude Online channel
field_reps = filtered[filtered["Sales Executive"] != "Online"]

# ─── Page Header ──────────────────────────────────────────────────────────────
st.markdown("""
<div class='page-title'> Rhythm Lonavala — Sales Productivity Dashboard</div>
<div class='page-sub'>Comprehensive overview of Room Nights, Revenue & ARR performance across your sales team</div>
""", unsafe_allow_html=True)

# ─── KPI Cards ────────────────────────────────────────────────────────────────
total_rn  = int(filtered["Room Nights"].sum())
total_rev = filtered["Revenue"].sum()
avg_arr   = filtered[filtered["ARR"] > 0]["ARR"].mean() if (filtered["ARR"] > 0).any() else 0
top_row   = field_reps.groupby("Sales Executive")["Revenue"].sum()
top_exec  = top_row.idxmax() if len(top_row) else "N/A"
top_rev   = top_row.max() if len(top_row) else 0
n_months  = filtered["Month-Year"].nunique()

k1, k2, k3, k4, k5 = st.columns(5)
for col, icon, label, val, sub in [
    (k1, "🛏️", "TOTAL ROOM NIGHTS", f"{total_rn:,}",   "All selected data"),
    (k2, "💰", "TOTAL REVENUE",      fmt_inr(total_rev), "Combined all channels"),
    (k3, "📈", "AVG. ARR",           fmt_inr(avg_arr),   "Average room rate"),
    (k4, "🏆", "TOP PERFORMER",       top_exec,           fmt_inr(top_rev)),
    (k5, "📅", "MONTHS IN VIEW",      str(n_months),      "Out of 10 total"),
]:
    with col:
        st.markdown(f"""
        <div class='kpi-card'>
            <div class='kpi-icon'>{icon}</div>
            <div class='kpi-label'>{label}</div>
            <div class='kpi-value'>{val}</div>
            <div class='kpi-sub'>{sub}</div>
        </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# CHART 1 — Monthly Revenue + Room Nights Trend (Full Width)
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("<div class='section-header'>📊 Monthly Total Revenue & Room Night Trend</div>", unsafe_allow_html=True)

monthly = (
    filtered.groupby(["Month-Year","Month_Order"])
    .agg(Revenue=("Revenue","sum"), Room_Nights=("Room Nights","sum"))
    .reset_index().sort_values("Month_Order")
)

fig1 = make_subplots(specs=[[{"secondary_y": True}]])
fig1.add_trace(go.Bar(
    x=monthly["Month-Year"], y=monthly["Revenue"],
    name="Revenue (₹)",
    marker=dict(color="#1e6fa8", line=dict(color="#0d3b66", width=0.5)),
    text=[fmt_inr(v) for v in monthly["Revenue"]],
    textposition="outside",
    textfont=dict(size=9, color="#0d3b66"),
    hovertemplate="<b>%{x}</b><br>Revenue: ₹%{y:,.0f}<extra></extra>",
), secondary_y=False)

fig1.add_trace(go.Scatter(
    x=monthly["Month-Year"], y=monthly["Room_Nights"],
    name="Room Nights",
    mode="lines+markers+text",
    line=dict(color="#e84393", width=2.5),
    marker=dict(size=8, color="#e84393"),
    text=monthly["Room_Nights"].astype(int),
    textposition="top center",
    textfont=dict(size=10, color="#8e44ad"),
    hovertemplate="<b>%{x}</b><br>Room Nights: %{y}<extra></extra>",
), secondary_y=True)

fig1.update_yaxes(title_text="Revenue (₹)", secondary_y=False, gridcolor=GRID_COLOR,
                  tickfont_size=10, title_font_color="#1e6fa8", tickformat=",", tickprefix="₹")
fig1.update_yaxes(title_text="Room Nights", secondary_y=True,
                  gridcolor="rgba(0,0,0,0)", title_font_color="#e84393", tickfont_size=10)
fig1.update_layout(
    plot_bgcolor=PLOT_BG, paper_bgcolor=PAPER_BG, font_color=FONT_COLOR,
    height=430, margin=dict(l=80, r=70, t=20, b=60),
    legend=dict(orientation="h", y=1.08, x=0.5, xanchor="center", font_size=12),
    xaxis=dict(gridcolor=GRID_COLOR, tickfont_size=12),
    hovermode="x unified", bargap=0.3,
)
st.plotly_chart(fig1, use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# CHART 2 — Executive Revenue Stacked by Month (Full Width)
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("<div class='section-header'>👥 Executive-wise Revenue by Month (Stacked)</div>", unsafe_allow_html=True)

exec_monthly = (
    filtered.groupby(["Sales Executive","Month-Year","Month_Order"])
    .agg(Revenue=("Revenue","sum"))
    .reset_index().sort_values("Month_Order")
)

fig2 = px.bar(
    exec_monthly, x="Month-Year", y="Revenue",
    color="Sales Executive", barmode="stack",
    color_discrete_sequence=COLORS, height=480,
    custom_data=["Sales Executive","Revenue"],
)
fig2.update_traces(
    hovertemplate="<b>%{customdata[0]}</b><br>Month: %{x}<br>Revenue: ₹%{customdata[1]:,.0f}<extra></extra>"
)
fig2.update_layout(
    plot_bgcolor=PLOT_BG, paper_bgcolor=PAPER_BG, font_color=FONT_COLOR,
    height=480, margin=dict(l=80, r=30, t=20, b=130),
    xaxis=dict(gridcolor=GRID_COLOR, tickfont_size=12),
    yaxis=dict(gridcolor=GRID_COLOR, tickfont_size=11, tickformat=",", tickprefix="₹"),
    legend=dict(orientation="h", y=-0.35, font_size=10, title_text=""),
    hovermode="x unified", bargap=0.25,
)
st.plotly_chart(fig2, use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# CHART 3 — Room Nights Grouped by Executive per Month (Full Width)
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("<div class='section-header'>🛏️ Executive-wise Room Nights by Month (Grouped)</div>", unsafe_allow_html=True)

rn_monthly = (
    filtered.groupby(["Sales Executive","Month-Year","Month_Order"])
    .agg(Room_Nights=("Room Nights","sum"))
    .reset_index().sort_values("Month_Order")
)

fig3 = px.bar(
    rn_monthly, x="Month-Year", y="Room_Nights",
    color="Sales Executive", barmode="group",
    color_discrete_sequence=COLORS, height=450,
    custom_data=["Sales Executive","Room_Nights"],
)
fig3.update_traces(
    hovertemplate="<b>%{customdata[0]}</b><br>Month: %{x}<br>Room Nights: %{customdata[1]:.0f}<extra></extra>"
)
fig3.update_layout(
    plot_bgcolor=PLOT_BG, paper_bgcolor=PAPER_BG, font_color=FONT_COLOR,
    height=450, margin=dict(l=60, r=30, t=20, b=130),
    xaxis=dict(gridcolor=GRID_COLOR, tickfont_size=12),
    yaxis=dict(gridcolor=GRID_COLOR, tickfont_size=11, title_text="Room Nights"),
    legend=dict(orientation="h", y=-0.35, font_size=10, title_text=""),
    hovermode="closest", bargap=0.15, bargroupgap=0.05,
)
st.plotly_chart(fig3, use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# CHART 4 & 5 — Revenue Pie | ARR Heatmap (Side by Side)
# ═══════════════════════════════════════════════════════════════════════════════
col4a, col4b = st.columns(2)

with col4a:
    st.markdown("<div class='section-header'>🥧 Revenue Share by Executive</div>", unsafe_allow_html=True)
    pie_data = filtered.groupby("Sales Executive")["Revenue"].sum().reset_index()
    pie_data = pie_data[pie_data["Revenue"] > 0].sort_values("Revenue", ascending=False)

    fig4 = go.Figure(go.Pie(
        labels=pie_data["Sales Executive"],
        values=pie_data["Revenue"],
        hole=0.45,
        marker_colors=COLORS[:len(pie_data)],
        textinfo="label+percent",
        textfont_size=10,
        hovertemplate="<b>%{label}</b><br>Revenue: ₹%{value:,.0f}<br>Share: %{percent}<extra></extra>",
        pull=[0.03] * len(pie_data),
    ))
    fig4.update_layout(
        plot_bgcolor=PLOT_BG, paper_bgcolor=PAPER_BG, font_color=FONT_COLOR,
        height=420, margin=dict(l=10, r=10, t=20, b=10),
        legend=dict(font_size=10, bgcolor="rgba(255,255,255,0.7)"),
        annotations=[dict(text="Revenue", x=0.5, y=0.5, font_size=13,
                          showarrow=False, font_color="#0d3b66")]
    )
    st.plotly_chart(fig4, use_container_width=True)

with col4b:
    st.markdown("<div class='section-header'>🌡️ ARR Heatmap — Executive × Month</div>", unsafe_allow_html=True)
    heat_df = filtered[filtered["ARR"] > 0].copy()
    if not heat_df.empty:
        pivot = heat_df.pivot_table(
            index="Sales Executive", columns="Month-Year", values="ARR", aggfunc="mean"
        )
        ordered_cols = [m for m in all_months if m in pivot.columns]
        pivot = pivot[ordered_cols]
        z_text = [[fmt_inr(v) if not pd.isna(v) else "" for v in row] for row in pivot.values]

        fig5 = go.Figure(go.Heatmap(
            z=pivot.values,
            x=pivot.columns.tolist(),
            y=pivot.index.tolist(),
            colorscale="Blues",
            text=z_text, texttemplate="%{text}",
            textfont=dict(size=8, color="#0d3b66"),
            hoverongaps=False,
            hovertemplate="<b>%{y}</b> | %{x}<br>ARR: ₹%{z:,.0f}<extra></extra>",
            colorbar=dict(title="ARR (₹)", tickfont=dict(color=FONT_COLOR, size=9),
                          title_font_color=FONT_COLOR, tickformat=","),
        ))
        fig5.update_layout(
            plot_bgcolor=PLOT_BG, paper_bgcolor=PAPER_BG, font_color=FONT_COLOR,
            height=420, margin=dict(l=140, r=20, t=20, b=80),
            xaxis=dict(tickangle=-30, tickfont_size=10),
            yaxis=dict(tickfont_size=9),
        )
        st.plotly_chart(fig5, use_container_width=True)
    else:
        st.info("No ARR data available for selected filters.")

# ═══════════════════════════════════════════════════════════════════════════════
# LEADERBOARD (Field Executives Only — excluding Online)
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("<div class='section-header'>🏆 Sales Leaderboard (Field Executives Only)</div>", unsafe_allow_html=True)

leaderboard = (
    field_reps.groupby("Sales Executive")
    .agg(Revenue=("Revenue","sum"), Room_Nights=("Room Nights","sum"), ARR=("ARR","mean"))
    .sort_values("Revenue", ascending=False).reset_index()
)
# Only show executives who have actual revenue
leaderboard = leaderboard[leaderboard["Revenue"] > 0]

medals   = ["🥇","🥈","🥉"] + [f"#{i}" for i in range(4, 30)]
total_lb = leaderboard["Revenue"].sum()
lb_c1, lb_c2 = st.columns(2)

for idx, (_, row) in enumerate(leaderboard.iterrows()):
    medal = medals[idx]
    pct   = (row["Revenue"] / total_lb * 100) if total_lb > 0 else 0
    col   = lb_c1 if idx % 2 == 0 else lb_c2
    with col:
        st.markdown(f"""
        <div class='rank-row'>
            <span style='font-size:1.3rem;min-width:40px;display:inline-block;'>{medal}</span>
            <div style='flex:1;padding:0 10px;'>
                <div style='color:#0d3b66;font-weight:700;font-size:0.92rem;'>{row['Sales Executive']}</div>
                <div style='color:#4a7fa5;font-size:0.78rem;margin-top:2px;'>
                    🛏️ {int(row['Room_Nights'])} Room Nights &nbsp;|&nbsp; ARR {fmt_inr(row['ARR'])}
                </div>
            </div>
            <div style='text-align:right;'>
                <div style='color:#1e6fa8;font-weight:700;font-size:1rem;'>{fmt_inr(row['Revenue'])}</div>
                <div style='color:#5b9bd5;font-size:0.75rem;'>{pct:.1f}% of total</div>
            </div>
        </div>""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# ONLINE vs FIELD COMPARISON (Full Width)
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("<div class='section-header'>🌐 Online vs Field Revenue Comparison by Month</div>", unsafe_allow_html=True)

comp_df = filtered.copy()
comp_df["Channel"] = comp_df["Sales Executive"].apply(lambda x: "Online" if x == "Online" else "Field Sales")
channel_monthly = (
    comp_df.groupby(["Channel","Month-Year","Month_Order"])
    .agg(Revenue=("Revenue","sum"))
    .reset_index().sort_values("Month_Order")
)

fig_comp = px.bar(
    channel_monthly, x="Month-Year", y="Revenue",
    color="Channel", barmode="group",
    color_discrete_map={"Online": "#27ae60", "Field Sales": "#1e6fa8"},
    height=380,
    custom_data=["Channel","Revenue"],
)
fig_comp.update_traces(
    hovertemplate="<b>%{customdata[0]}</b><br>%{x}<br>Revenue: ₹%{customdata[1]:,.0f}<extra></extra>"
)
fig_comp.update_layout(
    plot_bgcolor=PLOT_BG, paper_bgcolor=PAPER_BG, font_color=FONT_COLOR,
    height=380, margin=dict(l=80, r=30, t=20, b=60),
    xaxis=dict(gridcolor=GRID_COLOR, tickfont_size=12),
    yaxis=dict(gridcolor=GRID_COLOR, tickfont_size=11, tickformat=",", tickprefix="₹"),
    legend=dict(orientation="h", y=1.08, x=0.5, xanchor="center", font_size=12),
    hovermode="x unified", bargap=0.3,
)
st.plotly_chart(fig_comp, use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# INDIVIDUAL EXECUTIVE DEEP DIVE
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("<div class='section-header'>🔍 Individual Executive Deep Dive</div>", unsafe_allow_html=True)

exec_pick = st.selectbox(
    "Select an executive to analyse:",
    options=sorted(df["Sales Executive"].unique()),
    key="deepdive_exec"
)
exec_df = df[df["Sales Executive"] == exec_pick].sort_values("Month_Order")

# Revenue line — full width
fig_d1 = go.Figure()
fig_d1.add_trace(go.Scatter(
    x=exec_df["Month-Year"], y=exec_df["Revenue"],
    mode="lines+markers+text",
    line=dict(color="#1e6fa8", width=2.5),
    marker=dict(size=9, color="#1e6fa8"),
    text=[fmt_inr(v) for v in exec_df["Revenue"]],
    textposition="top center", textfont=dict(size=9, color="#0d3b66"),
    fill="tozeroy", fillcolor="rgba(30,111,168,0.08)",
    hovertemplate="<b>%{x}</b><br>Revenue: ₹%{y:,.0f}<extra></extra>",
))
fig_d1.update_layout(
    title=dict(text=f"📈 {exec_pick} — Monthly Revenue", font_color="#0d3b66", font_size=13),
    plot_bgcolor=PLOT_BG, paper_bgcolor=PAPER_BG, font_color=FONT_COLOR,
    height=320, margin=dict(l=80, r=20, t=50, b=60),
    xaxis=dict(gridcolor=GRID_COLOR, tickangle=-30, tickfont_size=11),
    yaxis=dict(gridcolor=GRID_COLOR, tickfont_size=10, tickformat=",", tickprefix="₹"),
)
st.plotly_chart(fig_d1, use_container_width=True)

# ARR + Room Nights side by side
dd1, dd2 = st.columns(2)

with dd1:
    fig_d2 = go.Figure()
    fig_d2.add_trace(go.Bar(
        x=exec_df["Month-Year"], y=exec_df["ARR"],
        marker=dict(color="#f39c12", line=dict(color="#e67e22", width=0.5)),
        text=[fmt_inr(v) for v in exec_df["ARR"]],
        textposition="outside", textfont=dict(size=9, color="#0d3b66"),
        hovertemplate="<b>%{x}</b><br>ARR: ₹%{y:,.0f}<extra></extra>",
    ))
    fig_d2.update_layout(
        title=dict(text=f"💛 {exec_pick} — Monthly ARR", font_color="#0d3b66", font_size=13),
        plot_bgcolor=PLOT_BG, paper_bgcolor=PAPER_BG, font_color=FONT_COLOR,
        height=320, margin=dict(l=80, r=20, t=50, b=60),
        xaxis=dict(gridcolor=GRID_COLOR, tickangle=-30, tickfont_size=10),
        yaxis=dict(gridcolor=GRID_COLOR, tickfont_size=10, tickformat=",", tickprefix="₹"),
        bargap=0.35,
    )
    st.plotly_chart(fig_d2, use_container_width=True)

with dd2:
    fig_d3 = go.Figure()
    fig_d3.add_trace(go.Bar(
        x=exec_df["Month-Year"], y=exec_df["Room Nights"],
        marker=dict(color="#27ae60", line=dict(color="#1e8449", width=0.5)),
        text=exec_df["Room Nights"].astype(int),
        textposition="outside", textfont=dict(size=10, color="#0d3b66"),
        hovertemplate="<b>%{x}</b><br>Room Nights: %{y}<extra></extra>",
    ))
    fig_d3.update_layout(
        title=dict(text=f"🛏️ {exec_pick} — Monthly Room Nights", font_color="#0d3b66", font_size=13),
        plot_bgcolor=PLOT_BG, paper_bgcolor=PAPER_BG, font_color=FONT_COLOR,
        height=320, margin=dict(l=50, r=20, t=50, b=60),
        xaxis=dict(gridcolor=GRID_COLOR, tickangle=-30, tickfont_size=10),
        yaxis=dict(gridcolor=GRID_COLOR, tickfont_size=11, title_text="Room Nights"),
        bargap=0.35,
    )
    st.plotly_chart(fig_d3, use_container_width=True)

# ─── Raw Data ─────────────────────────────────────────────────────────────────
with st.expander("📋 View Full Data Table"):
    disp = filtered.drop(columns=["Month_Order"]).copy()
    disp["Revenue"]     = disp["Revenue"].apply(fmt_inr)
    disp["ARR"]         = disp["ARR"].apply(fmt_inr)
    disp["Room Nights"] = disp["Room Nights"].astype(int)
    st.dataframe(disp, use_container_width=True, hide_index=True)

# ─── Footer ───────────────────────────────────────────────────────────────────
st.markdown("""
<hr style='border:none;border-top:2px solid #b3d9f5;margin-top:28px;'>
<div style='text-align:center;color:#4a7fa5;font-size:0.76rem;padding:10px 0;'>
    Rhythm Lonavala &nbsp;|&nbsp; Sales Productivity Dashboard &nbsp;
</div>
""", unsafe_allow_html=True)
