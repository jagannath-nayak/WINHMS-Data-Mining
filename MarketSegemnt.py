import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import io

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Market Segment Analysis",
    page_icon="🏨",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        font-size: 2rem; font-weight: 700; color: #1E3A5F;
        text-align: center; margin-bottom: 0.2rem;
    }
    .sub-header {
        text-align: center; color: #555; margin-bottom: 0.3rem;
        font-size: 1.05rem; font-weight: 600;
    }
    .date-range {
        text-align: center; color: #888; font-size: 0.85rem; margin-bottom: 1.5rem;
    }
    /* Make metric values never truncate */
    .kpi-card {
        background: #f4f7fb;
        border-radius: 10px;
        padding: 0.8rem 1rem;
        border-left: 4px solid #1E3A5F;
        min-width: 0;
    }
    .kpi-label {
        font-size: 0.78rem;
        color: #555;
        margin-bottom: 0.3rem;
        font-weight: 500;
    }
    .kpi-value {
        font-size: 1.1rem;
        font-weight: 700;
        color: #1E3A5F;
        word-break: break-all;
        white-space: normal;
        line-height: 1.3;
    }
</style>
""", unsafe_allow_html=True)


# ── Helper: parse the .xls / .xlsx ───────────────────────────────────────────
def parse_file(uploaded_file):
    suffix = os.path.splitext(uploaded_file.name)[-1].lower()
    uploaded_file.seek(0)
    buf = io.BytesIO(uploaded_file.read())

    try:
        engine = "xlrd" if suffix == ".xls" else "openpyxl"
        raw = pd.read_excel(buf, header=None, engine=engine)
    except ImportError:
        missing = "xlrd" if suffix == ".xls" else "openpyxl"
        st.error(f"Missing library: **{missing}**. Run: `pip install {missing}` then restart.")
        return None
    except Exception as e:
        st.error(f"Could not read file: {e}")
        return None

    # Detect header row (contains 'Company' and 'Nights')
    header_row = None
    for i, row in raw.iterrows():
        vals = [str(v).strip().lower() for v in row.values]
        if "company" in vals and "nights" in vals:
            header_row = i
            break

    if header_row is None:
        st.error("Could not detect the data header row.")
        return None

    # Extract property name and date range
    meta = {"property": "", "from": "", "to": ""}
    for i in range(header_row):
        row_vals = [str(v).strip() for v in raw.iloc[i].values if str(v).strip() not in ("", "nan")]
        for idx, v in enumerate(row_vals):
            if v.lower() == "from :" and idx + 1 < len(row_vals):
                meta["from"] = row_vals[idx + 1]
            elif v.lower() == "to :" and idx + 1 < len(row_vals):
                meta["to"] = row_vals[idx + 1]
            elif not meta["property"] and v.lower() not in ("from :", "to :"):
                meta["property"] = v

    # Rebuild from full raw scan for from/to
    for i in range(header_row):
        cells = list(raw.iloc[i].values)
        for j, val in enumerate(cells):
            s = str(val).strip()
            if "from" in s.lower() and ":" in s:
                # next non-nan cell
                for k in range(j+1, len(cells)):
                    nxt = str(cells[k]).strip()
                    if nxt not in ("", "nan"):
                        meta["from"] = nxt
                        break
            if s.lower() in ("to :", "to:"):
                for k in range(j+1, len(cells)):
                    nxt = str(cells[k]).strip()
                    if nxt not in ("", "nan"):
                        meta["to"] = nxt
                        break

    # Build dataframe from header row
    cols = [str(c).strip() for c in raw.iloc[header_row].values]
    data = raw.iloc[header_row + 1:].copy()
    data.columns = cols

    # Drop unnamed cols, blank rows, total row
    data = data.loc[:, ~data.columns.str.lower().str.startswith("unnamed")]
    data = data.dropna(subset=["Company"])
    data = data[~data["Company"].astype(str).str.strip().str.lower().isin(["nan", "total", ""])]

    # Coerce numerics
    num_cols = ["Nights", "Occupancy %", "Pax", "Room Revenue", "Revenue%", "ARR", "ARP"]
    for col in num_cols:
        if col in data.columns:
            data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0)

    data["Company"] = data["Company"].astype(str).str.strip()
    data = data[data["Nights"] > 0]

    return data, meta


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 Upload Data")
    uploaded = st.file_uploader(
        "Upload Market Segment File",
        type=["xls", "xlsx"],
        help="Upload the Market Segment Analysis export (.xls or .xlsx)",
    )
    st.markdown("---")
    st.markdown("### 🔍 Filters")

    df = None
    meta = {}
    df_filtered = None
    metric = "Room Revenue"

    if uploaded:
        result = parse_file(uploaded)
        if result:
            df, meta = result
            all_segments = sorted(df["Company"].unique().tolist())
            selected_segments = st.multiselect("Select Segments", all_segments, default=all_segments)
            metric = st.selectbox(
                "Primary Metric",
                ["Room Revenue", "Nights", "Pax", "ARR", "ARP", "Occupancy %", "Revenue%"],
            )
            df_filtered = df[df["Company"].isin(selected_segments)]

    st.markdown("---")
    st.caption("💡 pip install streamlit pandas plotly xlrd openpyxl")


# ── Main ──────────────────────────────────────────────────────────────────────
st.markdown('<div class="main-header">🏨 Market Segment Analysis Dashboard</div>', unsafe_allow_html=True)

if df is None or df_filtered is None:
    st.markdown('<div class="sub-header">Upload your Market Segment Analysis file from the sidebar.</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    c1.info("**Step 1** — Upload your `.xls` or `.xlsx` file from the sidebar.")
    c2.info("**Step 2** — Filter segments and choose a primary metric.")
    c3.info("**Step 3** — Explore KPIs, charts, and the data table.")
    st.stop()

# Property name + date range
prop = meta.get("property", "")
frm  = meta.get("from", "")
to   = meta.get("to", "")
st.markdown(f'<div class="sub-header">{prop}</div>', unsafe_allow_html=True)
if frm or to:
    st.markdown(f'<div class="date-range">📅 Period: {frm} &nbsp;→&nbsp; {to}</div>', unsafe_allow_html=True)

# ── KPI Cards ─────────────────────────────────────────────────────────────────
total_nights  = int(df_filtered["Nights"].sum())
total_revenue = df_filtered["Room Revenue"].sum()
total_pax     = int(df_filtered["Pax"].sum())
avg_arr       = df_filtered["ARR"].mean()
avg_arp       = df_filtered["ARP"].mean()
num_segments  = df_filtered["Company"].nunique()
total_occ     = df_filtered["Occupancy %"].sum()

kpis = [
    ("🛏️ Total Nights",  f"{total_nights:,}"),
    ("💰 Room Revenue",  f"Rs.{total_revenue:,.2f}"),
    ("👥 Total Pax",     f"{total_pax:,}"),
    ("📊 Avg ARR",       f"Rs.{avg_arr:,.2f}"),
    ("📊 Avg ARP",       f"Rs.{avg_arp:,.2f}"),
    ("🏷️ Segments",      f"{num_segments}"),
]
kpi_cols = st.columns(6)
for col, (label, value) in zip(kpi_cols, kpis):
    col.markdown(
        f'''<div class="kpi-card">
              <div class="kpi-label">{label}</div>
              <div class="kpi-value">{value}</div>
            </div>''',
        unsafe_allow_html=True,
    )

st.markdown("---")


# ── Utility: format revenue labels exactly (no M/K/B abbreviation) ────────────
def fmt_inr(val):
    """Return exact Indian-style number string e.g. 13,226,990.64"""
    return f"₹{val:,.2f}"


# ── Row 1: Bar + Pie ──────────────────────────────────────────────────────────
c1, c2 = st.columns(2)

with c1:
    st.subheader(f"📊 {metric} by Segment")

    plot_df = df_filtered.sort_values(metric, ascending=True).copy()

    # Build exact text labels — no abbreviation
    if metric == "Room Revenue":
        plot_df["_label"] = plot_df[metric].apply(lambda v: f"₹{v:,.2f}")
    elif metric in ("ARR", "ARP"):
        plot_df["_label"] = plot_df[metric].apply(lambda v: f"₹{v:,.2f}")
    elif metric == "Occupancy %":
        plot_df["_label"] = plot_df[metric].apply(lambda v: f"{v:.2f}%")
    elif metric == "Revenue%":
        plot_df["_label"] = plot_df[metric].apply(lambda v: f"{v:.2f}%")
    else:
        plot_df["_label"] = plot_df[metric].apply(lambda v: f"{int(v):,}")

    fig_bar = px.bar(
        plot_df,
        x=metric, y="Company",
        orientation="h",
        color=metric,
        color_continuous_scale="Blues",
        text="_label",
        height=480,
        hover_data={"_label": True, metric: False},
    )
    fig_bar.update_traces(textposition="outside", cliponaxis=False)
    fig_bar.update_layout(
        margin=dict(l=10, r=160, t=10, b=10),
        coloraxis_showscale=False,
        yaxis_title="",
        xaxis=dict(
            tickformat=",",          # comma-separated, no M/K
            separatethousands=True,
        ),
    )
    st.plotly_chart(fig_bar, use_container_width=True)

with c2:
    st.subheader("🥧 Revenue Share by Segment")
    fig_pie = px.pie(
        df_filtered,
        names="Company",
        values="Room Revenue",
        hole=0.38,
        height=480,
        hover_data={"Room Revenue": ":.2f"},
    )
    fig_pie.update_traces(
        textposition="inside",
        textinfo="percent+label",
        hovertemplate="<b>%{label}</b><br>Revenue: ₹%{value:,.2f}<br>Share: %{percent}<extra></extra>",
    )
    fig_pie.update_layout(margin=dict(l=10, r=10, t=10, b=10), showlegend=False)
    st.plotly_chart(fig_pie, use_container_width=True)


# ── Row 2: Scatter + ARR vs ARP ───────────────────────────────────────────────
c3, c4 = st.columns(2)

with c3:
    st.subheader("🔵 Nights vs Revenue (bubble = Pax)")
    fig_scatter = px.scatter(
        df_filtered,
        x="Nights", y="Room Revenue",
        size="Pax", color="Company",
        hover_name="Company",
        hover_data={"Room Revenue": ":.2f", "Nights": True, "Pax": True},
        size_max=50,
        height=400,
    )
    fig_scatter.update_traces(
        hovertemplate="<b>%{hovertext}</b><br>Nights: %{x:,}<br>Revenue: ₹%{y:,.2f}<br>Pax: %{marker.size:,}<extra></extra>"
    )
    fig_scatter.update_layout(
        margin=dict(l=10, r=10, t=10, b=10),
        showlegend=False,
        yaxis=dict(tickformat=",", separatethousands=True),
        xaxis=dict(tickformat=","),
    )
    st.plotly_chart(fig_scatter, use_container_width=True)

with c4:
    st.subheader("📈 ARR vs ARP by Segment")
    fig_arr_arp = go.Figure()
    fig_arr_arp.add_trace(go.Bar(
        name="ARR",
        x=df_filtered["Company"],
        y=df_filtered["ARR"],
        marker_color="#1E3A5F",
        text=[f"₹{v:,.2f}" for v in df_filtered["ARR"]],
        textposition="outside",
        hovertemplate="<b>%{x}</b><br>ARR: ₹%{y:,.2f}<extra></extra>",
    ))
    fig_arr_arp.add_trace(go.Bar(
        name="ARP",
        x=df_filtered["Company"],
        y=df_filtered["ARP"],
        marker_color="#2D9CDB",
        text=[f"₹{v:,.2f}" for v in df_filtered["ARP"]],
        textposition="outside",
        hovertemplate="<b>%{x}</b><br>ARP: ₹%{y:,.2f}<extra></extra>",
    ))
    fig_arr_arp.update_layout(
        barmode="group",
        height=400,
        xaxis_tickangle=-40,
        margin=dict(l=10, r=10, t=30, b=100),
        legend=dict(orientation="h", yanchor="bottom", y=1),
        yaxis=dict(title="Rate (₹)", tickformat=",", separatethousands=True),
    )
    st.plotly_chart(fig_arr_arp, use_container_width=True)


# ── Occupancy chart ───────────────────────────────────────────────────────────
st.subheader("📉 Occupancy % Contribution by Segment")
occ_df = df_filtered.sort_values("Occupancy %", ascending=False).copy()
fig_occ = px.bar(
    occ_df,
    x="Company", y="Occupancy %",
    color="Occupancy %",
    color_continuous_scale="Teal",
    text=[f"{v:.2f}%" for v in occ_df["Occupancy %"]],
    height=340,
    hover_data={"Occupancy %": ":.2f"},
)
fig_occ.update_traces(textposition="outside")
fig_occ.update_layout(
    margin=dict(l=10, r=10, t=10, b=100),
    coloraxis_showscale=False,
    xaxis_tickangle=-35,
    xaxis_title="",
    yaxis_title="Occupancy %",
)
st.plotly_chart(fig_occ, use_container_width=True)


# ── Revenue % Contribution ────────────────────────────────────────────────────
st.subheader("💹 Revenue % Contribution by Segment")
rev_df = df_filtered.sort_values("Revenue%", ascending=False).copy()
fig_rev = px.bar(
    rev_df,
    x="Company", y="Revenue%",
    color="Revenue%",
    color_continuous_scale="Oranges",
    text=[f"{v:.2f}%" for v in rev_df["Revenue%"]],
    height=340,
)
fig_rev.update_traces(textposition="outside")
fig_rev.update_layout(
    margin=dict(l=10, r=10, t=10, b=100),
    coloraxis_showscale=False,
    xaxis_tickangle=-35,
    xaxis_title="",
    yaxis_title="Revenue %",
)
st.plotly_chart(fig_rev, use_container_width=True)


# ── Data Table ────────────────────────────────────────────────────────────────
st.subheader("📋 Detailed Segment Data")

display_cols = ["Company", "Nights", "Occupancy %", "Pax", "Room Revenue", "Revenue%", "ARR", "ARP"]
display_cols = [c for c in display_cols if c in df_filtered.columns]

table_df = df_filtered[display_cols].sort_values("Room Revenue", ascending=False).reset_index(drop=True)

# Add totals row
totals = {
    "Company": "TOTAL",
    "Nights": int(table_df["Nights"].sum()),
    "Occupancy %": round(table_df["Occupancy %"].sum(), 2),
    "Pax": int(table_df["Pax"].sum()),
    "Room Revenue": round(table_df["Room Revenue"].sum(), 2),
    "Revenue%": round(table_df["Revenue%"].sum(), 2),
    "ARR": round(table_df["ARR"].mean(), 2),
    "ARP": round(table_df["ARP"].mean(), 2),
}
totals_row = pd.DataFrame([{c: totals.get(c, "") for c in display_cols}])
table_display = pd.concat([table_df, totals_row], ignore_index=True)

st.dataframe(
    table_display,
    use_container_width=True,
    height=420,
    column_config={
        "Room Revenue": st.column_config.NumberColumn("Room Revenue (₹)", format="₹%.2f"),
        "ARR":          st.column_config.NumberColumn("ARR (₹)",          format="₹%.2f"),
        "ARP":          st.column_config.NumberColumn("ARP (₹)",          format="₹%.2f"),
        "Nights":       st.column_config.NumberColumn("Nights",           format="%d"),
        "Pax":          st.column_config.NumberColumn("Pax",              format="%d"),
        "Occupancy %":  st.column_config.NumberColumn("Occupancy %",      format="%.2f%%"),
        "Revenue%":     st.column_config.NumberColumn("Revenue %",        format="%.2f%%"),
    },
)

# ── Download ──────────────────────────────────────────────────────────────────
csv = table_df.to_csv(index=False).encode()
st.download_button(
    "⬇️ Download Data as CSV",
    data=csv,
    file_name="market_segment_analysis.csv",
    mime="text/csv",
)