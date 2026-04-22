import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from datetime import datetime
import numpy as np

# ─── Page Config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Gtalk Adoption · IBCS Report",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── IBCS Color Palette & Constants ──────────────────────────────────────────
IBCS = {
    "AC":       "#1a1a1a",   # Actual – solid black
    "PY":       "#999999",   # Previous Year/Period – gray
    "PL":       "#ffffff",   # Plan – white (outlined)
    "FC":       "#666666",   # Forecast – hatched/medium gray
    "POS":      "#006400",   # Positive variance – dark green
    "NEG":      "#b30000",   # Negative variance – dark red
    "NEUTRAL":  "#999999",
    "BG":       "#ffffff",   # Background
    "GRID":     "#e0e0e0",   # Light grid
    "TEXT":     "#333333",   # Text
    "ACCENT":   "#2c5f8a",   # Subtle accent for highlights
}

# ─── IBCS CSS Theme ──────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* --- Global IBCS Theme --- */
    .stApp { background-color: #f7f7f7; }
    section[data-testid="stSidebar"] { background-color: #f0f0f0; border-right: 1px solid #ddd; }
    
    /* --- Typography --- */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; color: #333; }
    h1, h2, h3, h4 { font-family: 'Inter', sans-serif !important; color: #1a1a1a !important; font-weight: 600 !important; }
    
    /* --- Report Header --- */
    .report-header {
        background: white;
        border-bottom: 3px solid #1a1a1a;
        padding: 20px 30px;
        margin: -1rem -1rem 1.5rem -1rem;
    }
    .report-header h1 {
        font-size: 1.5rem !important;
        font-weight: 700 !important;
        margin: 0 !important;
        letter-spacing: -0.02em;
    }
    .report-header .subtitle {
        font-size: 0.85rem;
        color: #666;
        margin-top: 4px;
    }
    
    /* --- KPI Cards IBCS Style --- */
    .kpi-container {
        display: flex;
        gap: 16px;
        margin-bottom: 24px;
    }
    .kpi-card {
        background: white;
        border: 1px solid #ddd;
        border-top: 3px solid #1a1a1a;
        padding: 16px 20px;
        flex: 1;
        min-width: 0;
    }
    .kpi-card .kpi-label {
        font-size: 0.72rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #888;
        font-weight: 600;
        margin-bottom: 6px;
    }
    .kpi-card .kpi-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1a1a1a;
        line-height: 1.1;
    }
    .kpi-card .kpi-sub {
        font-size: 0.78rem;
        margin-top: 6px;
        color: #666;
    }
    .kpi-card .kpi-delta-pos { color: #006400; font-weight: 600; }
    .kpi-card .kpi-delta-neg { color: #b30000; font-weight: 600; }
    .kpi-card .kpi-delta-neutral { color: #999; font-weight: 500; }
    
    /* --- Section Container --- */
    .ibcs-section {
        background: white;
        border: 1px solid #ddd;
        padding: 20px 24px;
        margin-bottom: 20px;
    }
    .ibcs-section h3 {
        font-size: 0.95rem !important;
        font-weight: 600 !important;
        border-bottom: 2px solid #1a1a1a;
        padding-bottom: 8px;
        margin-bottom: 16px !important;
        text-transform: uppercase;
        letter-spacing: 0.04em;
    }
    .ibcs-section .section-msg {
        font-size: 0.82rem;
        color: #555;
        margin-top: -12px;
        margin-bottom: 14px;
        font-style: italic;
    }
    
    /* --- Legend --- */
    .ibcs-legend {
        display: flex;
        gap: 20px;
        align-items: center;
        font-size: 0.75rem;
        color: #666;
        margin-bottom: 12px;
    }
    .ibcs-legend .leg-item {
        display: flex;
        align-items: center;
        gap: 5px;
    }
    .leg-swatch {
        display: inline-block;
        width: 14px;
        height: 10px;
    }
    .leg-ac { background: #1a1a1a; }
    .leg-py { background: #999; }
    .leg-delta-pos { background: #006400; }
    .leg-delta-neg { background: #b30000; }
    
    /* --- Data Table IBCS --- */
    .ibcs-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.8rem;
    }
    .ibcs-table th {
        background: #f5f5f5;
        border-bottom: 2px solid #333;
        padding: 8px 10px;
        text-align: left;
        font-weight: 600;
        text-transform: uppercase;
        font-size: 0.72rem;
        letter-spacing: 0.04em;
        color: #555;
    }
    .ibcs-table th.num { text-align: right; }
    .ibcs-table td {
        padding: 7px 10px;
        border-bottom: 1px solid #eee;
        color: #333;
    }
    .ibcs-table td.num { text-align: right; font-variant-numeric: tabular-nums; }
    .ibcs-table tr:hover { background: #fafafa; }
    .ibcs-table .pos { color: #006400; font-weight: 600; }
    .ibcs-table .neg { color: #b30000; font-weight: 600; }
    .ibcs-table .neutral { color: #999999; font-weight: 500; }
    .ibcs-table .total-row td { border-top: 2px solid #333; font-weight: 700; background: #f9f9f9; }
    
    /* --- Mini bar in table --- */
    .mini-bar-bg {
        background: #eee;
        height: 6px;
        border-radius: 1px;
        position: relative;
        min-width: 80px;
    }
    .mini-bar-fill {
        height: 6px;
        border-radius: 1px;
        position: absolute;
        top: 0;
        left: 0;
    }
    
    /* --- Footer --- */
    .report-footer {
        font-size: 0.7rem;
        color: #aaa;
        text-align: center;
        padding: 16px;
        border-top: 1px solid #eee;
        margin-top: 20px;
    }
    
    /* Hide streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Selectbox/sidebar tweaks */
    .stSelectbox label, .stMultiSelect label {
        font-size: 0.8rem !important;
        font-weight: 600 !important;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: #666 !important;
    }
</style>
""", unsafe_allow_html=True)

# ─── IBCS Plotly Template ────────────────────────────────────────────────────
def ibcs_layout(**kwargs):
    """Return common IBCS layout dict for plotly figures."""
    base = dict(
        paper_bgcolor="white",
        plot_bgcolor="white",
        font=dict(family="Inter, sans-serif", size=11, color="#333"),
        margin=dict(l=50, r=20, t=40, b=40),
        xaxis=dict(
            showgrid=False,
            zeroline=True, zerolinecolor="#999", zerolinewidth=1,
            tickfont=dict(size=10, color="#666"),
        ),
        yaxis=dict(
            showgrid=True, gridcolor="#eee", gridwidth=1,
            zeroline=True, zerolinecolor="#999", zerolinewidth=1,
            tickfont=dict(size=10, color="#666"),
        ),
        showlegend=False,
        hoverlabel=dict(bgcolor="white", font_size=11, font_family="Inter"),
    )
    base.update(kwargs)
    return base


# ─── Data Loading ────────────────────────────────────────────────────────────
@st.cache_data(ttl=21600, show_spinner="Đang xử lý dữ liệu...")
def load_data():
    import urllib.request
    import io
    sheet_id = "1o2ODGcCfK9Y2flHztJCmliaNANIymvAx3QulzK6jet8"
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    res = urllib.request.urlopen(req)
    excel_data = res.read()
    
    df_staff = pd.read_excel(io.BytesIO(excel_data), sheet_name='data1')
    df_history = pd.read_excel(io.BytesIO(excel_data), sheet_name='data history')
    
    # Normalize date columns
    new_cols = []
    for c in df_history.columns:
        if isinstance(c, datetime):
            new_cols.append(c.strftime("%d/%m"))
        else:
            s = str(c)
            # Handle "DD/MM/YYYY" format → "DD/MM"
            if len(s) == 10 and '/' in s:
                parts = s.split('/')
                new_cols.append(f"{parts[0]}/{parts[1]}")
            elif len(s) >= 5 and '-' in s:
                # Handle "17-04"
                parts = s.split('-')
                if len(parts) >= 2:
                    new_cols.append(f"{parts[0]}/{parts[1]}")
                else:
                    new_cols.append(s)
            else:
                new_cols.append(s)
                
    df_history.columns = new_cols
    
    # Giữ nguyên int dtype để match chính xác
    # active_by_date: snapshot từng ngày, TOÀN BỘ ID (kể cả unmapped) dùng để đếm active
    active_by_date = {}
    for col in df_history.columns:
        ids = df_history[col].dropna()
        active_by_date[col] = set(ids.astype(int).unique())

    # KHÔNG concat unmapped vào df_staff — headcount chỉ tính từ data1
    return df_staff, active_by_date

df_staff, active_by_date = load_data()
all_dates = list(active_by_date.keys())

# ─── Sidebar Filters ────────────────────────────────────────────────────────
st.sidebar.markdown("""
<div style='border-bottom: 2px solid #333; padding-bottom: 10px; margin-bottom: 16px;'>
<strong style='font-size: 1rem; color: #1a1a1a;'>📊 BỘ LỌC BÁO CÁO</strong><br/>
<span style='font-size: 0.75rem; color: #888;'>Report Filters</span>
</div>
""", unsafe_allow_html=True)

if st.sidebar.button("🔄Reload data"):
    load_data.clear()
    st.rerun()

def safe_unique(s):
    return sorted([x for x in s.dropna().unique() if str(x).strip() != ""])

selected_date = st.sidebar.selectbox("NGÀY BÁO CÁO", options=all_dates, index=len(all_dates)-1)

# --- XỬ LÝ CLICK TỪ BIỂU ĐỒ (CROSS-FILTERING) ---
chart_selected_divisions = []
for chart_key in ["division_chart", "topbot_chart", "waterfall_chart"]:
    if chart_key in st.session_state:
        state = st.session_state[chart_key]
        if "selection" in state and "points" in state["selection"]:
            # Với biểu đồ ngang ngang (fig_div, fig_tb), tên division nằm ở pt["y"]
            # Với biểu đồ dọc (fig_wf), tên division nằm ở pt["x"]
            for pt in state["selection"]["points"]:
                if "y" in pt and isinstance(pt["y"], str):
                    chart_selected_divisions.append(pt["y"])
                elif "x" in pt and isinstance(pt["x"], str):
                    # Bỏ các dấu ... trong waterfall nếu có
                    val = pt["x"].replace("...", "")
                    # Tìm mapping tương đối (vì wf_names có thể bị cắt ngắn)
                    chart_selected_divisions.append(val)

all_divisions = safe_unique(df_staff['division_name_vn'])
selected_divisions = st.sidebar.multiselect("KHỐI (DIVISION)", all_divisions)

df_sidebar = df_staff.copy()
if selected_divisions:
    df_sidebar = df_sidebar[df_sidebar['division_name_vn'].isin(selected_divisions)]

all_departments = safe_unique(df_sidebar['department_name_vn'])
selected_departments = st.sidebar.multiselect("PHÒNG BAN (DEPARTMENT)", all_departments)
if selected_departments:
    df_sidebar = df_sidebar[df_sidebar['department_name_vn'].isin(selected_departments)]

all_sections = safe_unique(df_sidebar['section_name_vn'])
selected_sections = st.sidebar.multiselect("BỘ PHẬN (SECTION)", all_sections)
if selected_sections:
    df_sidebar = df_sidebar[df_sidebar['section_name_vn'].isin(selected_sections)]

all_teams = safe_unique(df_sidebar['team_name_vn'])
selected_teams = st.sidebar.multiselect("TEAM", all_teams)
if selected_teams:
    df_sidebar = df_sidebar[df_sidebar['team_name_vn'].isin(selected_teams)]

# Xử lý trường hợp waterfall cắt ngắn tên
if chart_selected_divisions:
    matched_divisions = []
    for cd in chart_selected_divisions:
        for ad in all_divisions:
            if cd.strip() in ad:
                matched_divisions.append(ad)
    chart_selected_divisions = list(set(chart_selected_divisions + matched_divisions))

df_filtered = df_sidebar.copy()
if chart_selected_divisions:
    df_filtered = df_filtered[df_filtered['division_name_vn'].isin(chart_selected_divisions)]


# Legend in sidebar
st.sidebar.markdown("---")
st.sidebar.markdown("""
<div style='font-size: 0.75rem; color: #888; line-height: 1.8;'>
<strong style="color:#555;">KÝ HIỆU IBCS</strong><br/>
<span style='display:inline-block;width:14px;height:10px;background:#1a1a1a;vertical-align:middle;margin-right:4px'></span> Hiện tại (AC)<br/>
<span style='display:inline-block;width:14px;height:10px;background:#999;vertical-align:middle;margin-right:4px'></span> Kỳ trước (PY)<br/>
<span style='display:inline-block;width:14px;height:10px;background:#006400;vertical-align:middle;margin-right:4px'></span> Chênh lệch dương (ΔPos)<br/>
<span style='display:inline-block;width:14px;height:10px;background:#b30000;vertical-align:middle;margin-right:4px'></span> Chênh lệch âm (ΔNeg)<br/>
</div>
""", unsafe_allow_html=True)


# ─── Metric Computation ─────────────────────────────────────────────────────
current_idx = all_dates.index(selected_date)
prev_date = all_dates[current_idx - 1] if current_idx > 0 else None
first_date = all_dates[0]

def compute_metrics(target_date, df):
    """
    total_hc      : Tổng nhân sự
    gtalk_all     : tổng ID trong snapshot ngày đó (kể cả unmapped)
    mapped_active : nhân sự trong df đang có trong Gtalk (intersection)
    pct_mapped    : mapped_active / total_hc * 100
    """
    if target_date is None:
        return 0, 0, 0, 0.0
    if len(df) == 0:
        return 0, 0, 0, 0.0
    active_set    = active_by_date[target_date]          # set of int
    total_hc      = df['employee_id'].nunique()
    gtalk_all     = len(active_set)
    # Convert về int để match đúng với active_set
    mapped_active = int(df['employee_id'].dropna().astype(int).isin(active_set).sum())
    pct_mapped    = (mapped_active / total_hc * 100) if total_hc > 0 else 0.0
    return total_hc, gtalk_all, mapped_active, pct_mapped

curr_total, curr_gtalk, curr_active, curr_pct = compute_metrics(selected_date, df_filtered)
prev_total, prev_gtalk, prev_active, prev_pct = compute_metrics(prev_date,     df_filtered)
first_total, first_gtalk, first_active, first_pct = compute_metrics(first_date, df_filtered)

inactive_count    = curr_total - curr_active
delta_active      = curr_active - prev_active
delta_gtalk       = curr_gtalk  - prev_gtalk
delta_pct         = curr_pct    - prev_pct
cumulative_growth = curr_pct    - first_pct


# ═══════════════════════════════════════════════════════════════════════════
# REPORT HEADER
# ═══════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="report-header">
<h1>GTALK WORKFORCE ADOPTION REPORT</h1>
<div class="subtitle">
Báo cáo tỷ lệ triển khai Gtalk · Ngày báo cáo: <strong>{selected_date}</strong> · 
Nhân sự đã active: <strong>{curr_active:,}/{curr_total:,}</strong> · Tổng Gtalk: <strong>{curr_gtalk:,}</strong> · Tỷ lệ: <strong>{curr_pct:.1f}%</strong>
</div>
</div>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
# KPI CARDS
# ═══════════════════════════════════════════════════════════════════════════
def delta_html(val, fmt="+,.0f", is_pct=False):
    suffix = "%" if is_pct else ""
    if val > 0:
        return f'<span class="kpi-delta-pos">▲ {val:{fmt}}{suffix}</span>'
    elif val < 0:
        return f'<span class="kpi-delta-neg">▼ {abs(val):{fmt}}{suffix}</span>'
    return f'<span class="kpi-delta-neutral">— 0{suffix}</span>'

pct_fmt = "+.1f"

st.markdown(f"""
<div class="kpi-container">

<div class="kpi-card">
<div class="kpi-label">👥 Tổng Nhân Sự</div>
<div class="kpi-value">{curr_total:,}</div>
<div class="kpi-sub">Tổng nhân sự </div>
</div>

<div class="kpi-card" style="border-top-color:#006400;">
<div class="kpi-label" style="color:#006400;">✅ Nhân Sự Đã Active</div>
<div class="kpi-value" style="color:#006400;">{curr_active:,}</div>
<div class="kpi-sub">Nhân sự trong biên chế đang dùng Gtalk · {delta_html(delta_active)} so kỳ trước</div>
</div>

<div class="kpi-card" style="border-top-color:#b30000;">
<div class="kpi-label" style="color:#b30000;">⏳ Chưa Active</div>
<div class="kpi-value" style="color:#b30000;">{inactive_count:,}</div>
<div class="kpi-sub">Nhân sự trong biên chế chưa có trong Gtalk ngày {selected_date}</div>
</div>

<div class="kpi-card" style="border-top-color:#2c5f8a;">
<div class="kpi-label" style="color:#2c5f8a;">📱 Đang Dùng Gtalk</div>
<div class="kpi-value" style="color:#2c5f8a;">{curr_gtalk:,}</div>
<div class="kpi-sub">Tổng user có trong Gtalk ngày {selected_date} · kể cả ngoài biên chế · {delta_html(delta_gtalk)} so kỳ trước</div>
</div>

<div class="kpi-card">
<div class="kpi-label">📈 Tỷ Lệ Adoption</div>
<div class="kpi-value">{curr_pct:.1f}%</div>
<div class="kpi-sub">Nhân sự đã active / Tổng HC · Δ kỳ trước: {delta_html(delta_pct, pct_fmt, True)} · Tích lũy: {cumulative_growth:+.1f}pp</div>
</div>

</div>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 1: HISTORICAL TREND (IBCS Line + Bar Combo)
# ═══════════════════════════════════════════════════════════════════════════
st.markdown('<div class="ibcs-section">', unsafe_allow_html=True)
st.markdown("<h3>XU HƯỚNG ADOPTION THEO THỜI GIAN</h3>", unsafe_allow_html=True)

trend_data = []
for i, d in enumerate(all_dates):
    _, _, active, pct = compute_metrics(d, df_filtered)
    prev_d = all_dates[i-1] if i > 0 else None
    _, _, prev_a, _ = compute_metrics(prev_d, df_filtered)
    new_users = active - prev_a if prev_d else 0
    trend_data.append({"date": d, "active": active, "pct": round(pct, 2), "new_users": new_users})

df_trend = pd.DataFrame(trend_data)

# Message line
if len(df_trend) >= 2:
    trend_msg = f"Tỷ lệ adoption tăng từ {df_trend['pct'].iloc[0]:.1f}% lên {df_trend['pct'].iloc[-1]:.1f}% (+{df_trend['pct'].iloc[-1] - df_trend['pct'].iloc[0]:.1f}pp) trong {len(all_dates)} kỳ báo cáo"
else:
    trend_msg = ""
st.markdown(f'<p class="section-msg">{trend_msg}</p>', unsafe_allow_html=True)

# Dual axis chart: bars for new users, line for %
fig_trend = make_subplots(specs=[[{"secondary_y": True}]])

# Bar: New active users per period (IBCS – gray bars)
fig_trend.add_trace(
    go.Bar(
        x=df_trend["date"], y=df_trend["new_users"],
        name="User mới",
        marker_color=IBCS["PY"],
        marker_line=dict(width=0),
        opacity=0.6,
        hovertemplate="<b>%{x}</b><br>User mới: %{y:,}<extra></extra>"
    ),
    secondary_y=False
)

# Line: Adoption % (IBCS – solid black)
fig_trend.add_trace(
    go.Scatter(
        x=df_trend["date"], y=df_trend["pct"],
        name="% Adoption",
        mode="lines+markers+text",
        line=dict(color=IBCS["AC"], width=2.5),
        marker=dict(size=5, color=IBCS["AC"]),
        text=[f"{v:.1f}%" for v in df_trend["pct"]],
        textposition="top center",
        textfont=dict(size=9, color="#333"),
        hovertemplate="<b>%{x}</b><br>Adoption: %{y:.1f}%<extra></extra>"
    ),
    secondary_y=True
)

fig_trend.update_layout(
    **ibcs_layout(
        height=340,
        margin=dict(l=50, r=50, t=20, b=50),
        showlegend=True,
    ),
    legend=dict(
        orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
        font=dict(size=10), bgcolor="rgba(0,0,0,0)",
    ),
    bargap=0.3,
)
fig_trend.update_xaxes(type="category", tickangle=-45)
fig_trend.update_yaxes(title_text="User mới (người)", secondary_y=False, showgrid=False)
fig_trend.update_yaxes(title_text="Tỷ lệ Adoption (%)", secondary_y=True, range=[0, 105], showgrid=True, gridcolor="#eee")

st.plotly_chart(fig_trend, use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 2: DIVISION BREAKDOWN – AC vs PY with Variance (IBCS Horizontal Bar)
# ═══════════════════════════════════════════════════════════════════════════
st.markdown('<div class="ibcs-section">', unsafe_allow_html=True)
st.markdown("<h3>PHÂN TÍCH THEO KHỐI (DIVISION) · TỶ LỆ ADOPTION</h3>", unsafe_allow_html=True)

if len(df_sidebar) > 0:
    active_set_curr = active_by_date[selected_date]
    active_set_prev = active_by_date[prev_date] if prev_date else set()
    
    # 1. DỮ LIỆU BIỂU ĐỒ TRÁI (LUÔN HIỂN THỊ CÁC KHỐI THEO SIDEBAR)
    df_chart = df_sidebar.copy()
    df_chart["is_active_curr"] = df_chart["employee_id"].isin(active_set_curr)
    df_chart["is_active_prev"] = df_chart["employee_id"].isin(active_set_prev)
    
    div_df = df_chart.groupby("division_name_vn").agg(
        total=("employee_id", "count"),
        active_curr=("is_active_curr", "sum"),
        active_prev=("is_active_prev", "sum"),
    ).reset_index()
    
    div_df["pct_curr"] = (div_df["active_curr"] / div_df["total"] * 100).round(1)
    div_df["pct_prev"] = (div_df["active_prev"] / div_df["total"] * 100).round(1)
    div_df["delta_pct"] = (div_df["pct_curr"] - div_df["pct_prev"]).round(1)
    div_df["delta_abs"] = div_df["active_curr"] - div_df["active_prev"]
    div_df = div_df.sort_values("pct_curr", ascending=True).reset_index(drop=True)
    
    div_df_display = div_df[div_df["total"] >= 3].copy()
    
    # 2. DỮ LIỆU BẢNG PHẢI (DRILL-DOWN VÀO PHÒNG BAN NẾU CHỌN 1 KHỐI)
    is_drilldown = len(chart_selected_divisions) == 1
    drill_col = "department_name_vn" if is_drilldown else "division_name_vn"
    drill_title = "Phòng Ban" if is_drilldown else "Khối"
    
    df_table = df_filtered.copy() if is_drilldown else df_sidebar.copy()
    df_table["is_active_curr"] = df_table["employee_id"].isin(active_set_curr)
    df_table["is_active_prev"] = df_table["employee_id"].isin(active_set_prev)
    
    table_df = df_table.groupby(drill_col).agg(
        total=("employee_id", "count"),
        active_curr=("is_active_curr", "sum"),
        active_prev=("is_active_prev", "sum"),
    ).reset_index()
    table_df = table_df.rename(columns={drill_col: "name"})
    
    table_df["pct_curr"] = (table_df["active_curr"] / table_df["total"] * 100).round(1)
    table_df["pct_prev"] = (table_df["active_prev"] / table_df["total"] * 100).round(1)
    table_df["delta_pct"] = (table_df["pct_curr"] - table_df["pct_prev"]).round(1)
    table_df["delta_abs"] = table_df["active_curr"] - table_df["active_prev"]
    
    if is_drilldown:
        drill_msg = f'<span style="color:#2c5f8a; font-weight:600;">Đang xem chi tiết {len(table_df)} Phòng Ban thuộc Khối: {chart_selected_divisions[0]}</span>'
    else:
        if len(div_df_display) > 0:
            msg_top = div_df_display.sort_values("pct_curr", ascending=False).iloc[0]
            msg_bottom = div_df_display.sort_values("pct_curr", ascending=True).iloc[0]
            drill_msg = f'Khối cao nhất: <b>{msg_top["division_name_vn"]}</b> ({msg_top["pct_curr"]:.1f}%) · Thấp nhất: <b>{msg_bottom["division_name_vn"]}</b> ({msg_bottom["pct_curr"]:.1f}%)'
        else:
            drill_msg = ""
            
    st.markdown(f'<p class="section-msg">{drill_msg}</p>', unsafe_allow_html=True)
    
    col_chart, col_table = st.columns([3, 2])
    
    with col_chart:
        # IBCS horizontal grouped bar: AC (black) vs PY (gray)
        fig_div = go.Figure()
        
        fig_div.add_trace(go.Bar(
            y=div_df_display["division_name_vn"],
            x=div_df_display["pct_prev"],
            orientation='h',
            name=f"Kỳ trước ({prev_date})" if prev_date else "N/A",
            marker_color=IBCS["PY"],
            marker_line=dict(width=0),
            text=[f"{v:.1f}%" for v in div_df_display["pct_prev"]],
            textposition="outside",
            textfont=dict(size=9, color="#999"),
            hovertemplate="<b>%{y}</b><br>Kỳ trước: %{x:.1f}%<extra></extra>",
        ))
        
        fig_div.add_trace(go.Bar(
            y=div_df_display["division_name_vn"],
            x=div_df_display["pct_curr"],
            orientation='h',
            name=f"Hiện tại ({selected_date})",
            marker_color=IBCS["AC"],
            marker_line=dict(width=0),
            text=[f"{v:.1f}%" for v in div_df_display["pct_curr"]],
            textposition="outside",
            textfont=dict(size=9, color="#333"),
            hovertemplate="<b>%{y}</b><br>Hiện tại: %{x:.1f}%<extra></extra>",
        ))
        
        fig_div.update_layout(
            **ibcs_layout(
                height=max(300, len(div_df_display) * 50),
                margin=dict(l=200, r=60, t=10, b=30),
                barmode='group',
                bargap=0.25,
                bargroupgap=0.05,
                showlegend=True,
            ),
            legend=dict(
                orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                font=dict(size=10),
            ),
        )
        fig_div.update_xaxes(range=[0, 105], dtick=25, showgrid=True, gridcolor="#eee")
        fig_div.update_yaxes(showgrid=False, tickfont=dict(size=10))
        
        st.plotly_chart(
            fig_div, 
            use_container_width=True, 
            on_select="rerun",
            selection_mode="points",
            key="division_chart"
        )
    
    with col_table:
        # IBCS structured table with inline variance
        rows_html = ""
        sorted_table = table_df.sort_values("pct_curr", ascending=False)
        for _, r in sorted_table.iterrows():
            if r["delta_pct"] > 0:
                delta_cls, delta_sign = "pos", "+"
            elif r["delta_pct"] < 0:
                delta_cls, delta_sign = "neg", ""
            else:
                delta_cls, delta_sign = "neutral", ""
            bar_width = min(r["pct_curr"], 100)
            bar_color = IBCS["AC"]
            
            # GIẢI PHÁP: Xóa khoảng trắng đầu dòng cho thẻ HTML để tránh lỗi Markdown
            rows_html += f"""
<tr>
<td style="max-width:160px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">{r['name']}</td>
<td class="num">{int(r['total']):,}</td>
<td class="num">{int(r['active_curr']):,}</td>
<td class="num" style="font-weight:600">{r['pct_curr']:.1f}%</td>
<td>
<div class="mini-bar-bg">
<div class="mini-bar-fill" style="width:{bar_width}%;background:{bar_color};"></div>
</div>
</td>
<td class="num {delta_cls}">{delta_sign}{r['delta_pct']:.1f}pp</td>
<td class="num {delta_cls}">{delta_sign}{int(r['delta_abs']):,}</td>
</tr>
"""
        
        # Totals
        t_total = sorted_table['total'].sum()
        t_active = sorted_table['active_curr'].sum()
        t_prev = sorted_table['active_prev'].sum()
        t_pct = (t_active/t_total*100) if t_total > 0 else 0
        t_pct_prev = (t_prev/t_total*100) if t_total > 0 else 0
        t_delta_pct = t_pct - t_pct_prev
        t_delta_abs = t_active - t_prev
        if t_delta_pct > 0:
            t_delta_cls, t_delta_sign = "pos", "+"
        elif t_delta_pct < 0:
            t_delta_cls, t_delta_sign = "neg", ""
        else:
            t_delta_cls, t_delta_sign = "neutral", ""
        
        # GIẢI PHÁP: Xóa khoảng trắng đầu dòng
        rows_html += f"""
<tr class="total-row">
<td><strong>TỔNG</strong></td>
<td class="num">{int(t_total):,}</td>
<td class="num">{int(t_active):,}</td>
<td class="num" style="font-weight:700">{t_pct:.1f}%</td>
<td></td>
<td class="num {t_delta_cls}">{t_delta_sign}{t_delta_pct:.1f}pp</td>
<td class="num {t_delta_cls}">{t_delta_sign}{int(t_delta_abs):,}</td>
</tr>
"""
        
        # GIẢI PHÁP: Xóa khoảng trắng đầu dòng
        st.markdown(f"""
<table class="ibcs-table">
<thead>
<tr>
<th>{drill_title}</th>
<th class="num">HC</th>
<th class="num">Active</th>
<th class="num">%</th>
<th style="min-width:80px"></th>
<th class="num">Δ%</th>
<th class="num">Δ Abs</th>
</tr>
</thead>
<tbody>{rows_html}</tbody>
</table>
""", unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 3: VARIANCE WATERFALL & TOP/BOTTOM ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════
col_wf, col_topbot = st.columns(2)

with col_wf:
    st.markdown('<div class="ibcs-section">', unsafe_allow_html=True)
    st.markdown("<h3>WATERFALL: THAY ĐỔI SỐ LƯỢNG ACTIVE</h3>", unsafe_allow_html=True)
    
    if len(df_filtered) > 0 and prev_date:
        # Compute change by division
        wf_data = div_df_display.sort_values("delta_abs", ascending=False).copy()
        wf_data = wf_data[wf_data["delta_abs"] != 0]
        
        if len(wf_data) > 0:
            st.markdown(
                f'<p class="section-msg">Thay đổi tổng: {delta_sign if delta_active >= 0 else ""}{delta_active:,} user từ {prev_date} → {selected_date}</p>',
                unsafe_allow_html=True
            )
            
            fig_wf = go.Figure()
            
            # Build waterfall
            wf_names = list(wf_data["division_name_vn"])
            wf_values = list(wf_data["delta_abs"].astype(int))
            
            colors = [IBCS["POS"] if v >= 0 else IBCS["NEG"] for v in wf_values]
            
            fig_wf.add_trace(go.Waterfall(
                orientation="v",
                x=[n[:20] + "..." if len(n) > 20 else n for n in wf_names],
                y=wf_values,
                connector=dict(line=dict(color="#ddd", width=1)),
                increasing=dict(marker=dict(color=IBCS["POS"])),
                decreasing=dict(marker=dict(color=IBCS["NEG"])),
                totals=dict(marker=dict(color=IBCS["AC"])),
                text=[f"{v:+,}" for v in wf_values],
                textposition="outside",
                textfont=dict(size=9),
                hovertemplate="<b>%{x}</b><br>Thay đổi: %{y:+,} user<extra></extra>",
            ))
            
            fig_wf.update_layout(
                **ibcs_layout(
                    height=340,
                    margin=dict(l=50, r=20, t=10, b=80),
                ),
            )
            fig_wf.update_xaxes(tickangle=-45, tickfont=dict(size=8))
            
            st.plotly_chart(
                fig_wf, 
                use_container_width=True,
                on_select="rerun",
                selection_mode="points",
                key="waterfall_chart"
            )
        else:
            st.info("Không có thay đổi giữa hai kỳ.")
    else:
        st.info("Chưa có dữ liệu kỳ trước để so sánh.")
    
    st.markdown('</div>', unsafe_allow_html=True)


with col_topbot:
    st.markdown('<div class="ibcs-section">', unsafe_allow_html=True)
    st.markdown("<h3>TOP & BOTTOM PERFORMERS</h3>", unsafe_allow_html=True)
    
    if len(df_filtered) > 0 and 'div_df_display' in dir():
        ranked = div_df_display[div_df_display["total"] >= 3].sort_values("pct_curr", ascending=False)
        
        if len(ranked) >= 2:
            top_5 = ranked.head(5)
            bottom_5 = ranked.tail(5)
            
            st.markdown('<p class="section-msg">Top 5 và Bottom 5 Khối theo tỷ lệ adoption</p>', unsafe_allow_html=True)
            
            fig_tb = go.Figure()
            
            # Bottom first (so they appear at bottom)
            combined = pd.concat([bottom_5.sort_values("pct_curr", ascending=True), 
                                  top_5.sort_values("pct_curr", ascending=True)])
            combined = combined.drop_duplicates(subset=["division_name_vn"])
            
            bar_colors = []
            for _, r in combined.iterrows():
                if r["division_name_vn"] in top_5["division_name_vn"].values:
                    bar_colors.append(IBCS["AC"])
                else:
                    bar_colors.append(IBCS["PY"])
            
            fig_tb.add_trace(go.Bar(
                y=combined["division_name_vn"],
                x=combined["pct_curr"],
                orientation='h',
                marker_color=bar_colors,
                text=[f"{v:.1f}%" for v in combined["pct_curr"]],
                textposition="outside",
                textfont=dict(size=9),
                hovertemplate="<b>%{y}</b><br>Adoption: %{x:.1f}%<extra></extra>",
            ))
            
            fig_tb.update_layout(
                **ibcs_layout(
                    height=340,
                    margin=dict(l=200, r=60, t=10, b=30),
                ),
            )
            fig_tb.update_xaxes(range=[0, 105], dtick=25, showgrid=True, gridcolor="#eee")
            fig_tb.update_yaxes(showgrid=False, tickfont=dict(size=9))
            
            st.plotly_chart(
                fig_tb, 
                use_container_width=True,
                on_select="rerun",
                selection_mode="points",
                key="topbot_chart"
            )
        else:
            st.info("Cần ít nhất 2 khối để so sánh.")
    
    st.markdown('</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 4: DETAILED BREAKDOWN TABLE (Drill-down by Department/Section/Team)
# ═══════════════════════════════════════════════════════════════════════════
st.markdown('<div class="ibcs-section">', unsafe_allow_html=True)
st.markdown("<h3>CHI TIẾT PHÂN BỔ (DRILL-DOWN TABLE)</h3>", unsafe_allow_html=True)

breakdown_level = st.selectbox(
    "Phân tích theo cấp:",
    ["Khối (Division)", "Phòng Ban (Department)", "Bộ Phận (Section)", "Team"],
    label_visibility="collapsed"
)
col_map = {
    "Khối (Division)": ["division_name_vn"],
    "Phòng Ban (Department)": ["division_name_vn", "department_name_vn"],
    "Bộ Phận (Section)": ["division_name_vn", "department_name_vn", "section_name_vn"],
    "Team": ["division_name_vn", "department_name_vn", "section_name_vn", "team_name_vn"]
}
group_cols = col_map[breakdown_level]

if len(df_filtered) > 0:
    active_set_c = active_by_date[selected_date]
    active_set_p = active_by_date[prev_date] if prev_date else set()
    
    df_bd = df_filtered.copy()
    df_bd["is_active_c"] = df_bd["employee_id"].isin(active_set_c)
    df_bd["is_active_p"] = df_bd["employee_id"].isin(active_set_p)
    
    bd_df = df_bd.groupby(group_cols, dropna=False).agg(
        total=("employee_id", "count"),
        active_c=("is_active_c", "sum"),
        active_p=("is_active_p", "sum"),
    ).reset_index()
    
    bd_df["pct_c"] = (bd_df["active_c"] / bd_df["total"] * 100).round(1)
    bd_df["pct_p"] = (bd_df["active_p"] / bd_df["total"] * 100).round(1)
    bd_df["delta_pct"] = (bd_df["pct_c"] - bd_df["pct_p"]).round(1)
    bd_df["delta_abs"] = bd_df["active_c"] - bd_df["active_p"]
    bd_df["inactive"] = bd_df["total"] - bd_df["active_c"]
    bd_df = bd_df.sort_values("pct_c", ascending=False).reset_index(drop=True)
    
    # Build HTML table
    bd_rows = ""
    for idx, r in bd_df.iterrows():
        if r["delta_pct"] > 0:
            d_cls, d_sign = "pos", "+"
        elif r["delta_pct"] < 0:
            d_cls, d_sign = "neg", ""
        else:
            d_cls, d_sign = "neutral", ""
        bar_w = min(r["pct_c"], 100)
        rank = idx + 1
        
        name_parts = []
        for c in group_cols:
            val = r[c]
            if pd.notna(val) and str(val).strip() != "":
                name_parts.append(str(val))
        disp_name = " › ".join(name_parts)
        if not disp_name:
            disp_name = "Chưa xác định"
        
        # GIẢI PHÁP: Xóa khoảng trắng đầu dòng
        bd_rows += f"""
<tr>
<td class="num" style="color:#aaa;font-size:0.7rem">{rank}</td>
<td style="max-width:350px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="{disp_name}">{disp_name}</td>
<td class="num">{int(r['total']):,}</td>
<td class="num">{int(r['active_c']):,}</td>
<td class="num" style="color:#888">{int(r['inactive']):,}</td>
<td class="num" style="font-weight:600">{r['pct_c']:.1f}%</td>
<td>
<div class="mini-bar-bg">
<div class="mini-bar-fill" style="width:{bar_w}%;background:{IBCS['AC']};"></div>
</div>
</td>
<td class="num" style="color:#888">{r['pct_p']:.1f}%</td>
<td class="num {d_cls}">{d_sign}{r['delta_pct']:.1f}pp</td>
<td class="num {d_cls}">{d_sign}{int(r['delta_abs']):,}</td>
</tr>
"""
    
    # Total row
    bd_t_total = bd_df['total'].sum()
    bd_t_ac = int(bd_df['active_c'].sum())
    bd_t_prev = int(bd_df['active_p'].sum())
    bd_t_inac = bd_t_total - bd_t_ac
    bd_t_pct = (bd_t_ac/bd_t_total*100) if bd_t_total > 0 else 0
    bd_t_pct_p = (bd_t_prev/bd_t_total*100) if bd_t_total > 0 else 0
    bd_t_delta_pct = bd_t_pct - bd_t_pct_p
    bd_t_delta_abs = bd_t_ac - bd_t_prev
    if bd_t_delta_pct > 0:
        bd_t_cls, bd_t_sign = "pos", "+"
    elif bd_t_delta_pct < 0:
        bd_t_cls, bd_t_sign = "neg", ""
    else:
        bd_t_cls, bd_t_sign = "neutral", ""
    
    # GIẢI PHÁP: Xóa khoảng trắng đầu dòng
    bd_rows += f"""
<tr class="total-row">
<td></td>
<td><strong>TỔNG CỘNG</strong></td>
<td class="num">{int(bd_t_total):,}</td>
<td class="num">{bd_t_ac:,}</td>
<td class="num" style="color:#888">{bd_t_inac:,}</td>
<td class="num" style="font-weight:700">{bd_t_pct:.1f}%</td>
<td></td>
<td class="num" style="color:#888">{bd_t_pct_p:.1f}%</td>
<td class="num {bd_t_cls}">{bd_t_sign}{bd_t_delta_pct:.1f}pp</td>
<td class="num {bd_t_cls}">{bd_t_sign}{bd_t_delta_abs:,}</td>
</tr>
"""
    
    # GIẢI PHÁP: Xóa khoảng trắng đầu dòng
    st.markdown(f"""
<div style="overflow-x:auto;">
<table class="ibcs-table">
<thead>
<tr>
<th class="num" style="width:30px">#</th>
<th>{breakdown_level}</th>
<th class="num">HC</th>
<th class="num">Active</th>
<th class="num">Chưa</th>
<th class="num">% AC</th>
<th style="min-width:80px"></th>
<th class="num">% PY</th>
<th class="num">Δ%</th>
<th class="num">Δ Abs</th>
</tr>
</thead>
<tbody>{bd_rows}</tbody>
</table>
</div>
""", unsafe_allow_html=True)
else:
    st.info("Không có dữ liệu cho bộ lọc hiện tại.")

st.markdown('</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 5: BU-LEVEL ANALYSIS & CUMULATIVE GROWTH
# ═══════════════════════════════════════════════════════════════════════════
col_bu, col_cum = st.columns(2)

with col_bu:
    st.markdown('<div class="ibcs-section">', unsafe_allow_html=True)
    st.markdown("<h3>PHÂN TÍCH THEO BU (BUSINESS UNIT)</h3>", unsafe_allow_html=True)
    
    if len(df_filtered) > 0 and "bu_name" in df_filtered.columns:
        df_bu_temp = df_filtered.copy()
        df_bu_temp["is_active"] = df_bu_temp["employee_id"].isin(active_by_date[selected_date])
        
        bu_df = df_bu_temp.groupby("bu_name").agg(
            total=("employee_id", "count"),
            active=("is_active", "sum"),
        ).reset_index()
        bu_df["pct"] = (bu_df["active"] / bu_df["total"] * 100).round(1)
        bu_df = bu_df.sort_values("total", ascending=False)
        
        # Stacked horizontal bar: active vs inactive
        fig_bu = go.Figure()
        
        fig_bu.add_trace(go.Bar(
            y=bu_df["bu_name"],
            x=bu_df["active"],
            orientation='h',
            name="Active",
            marker_color=IBCS["AC"],
            text=[f"{int(v):,}" for v in bu_df["active"]],
            textposition="inside",
            textfont=dict(size=9, color="white"),
            hovertemplate="<b>%{y}</b><br>Active: %{x:,}<extra></extra>",
        ))
        
        fig_bu.add_trace(go.Bar(
            y=bu_df["bu_name"],
            x=bu_df["total"] - bu_df["active"],
            orientation='h',
            name="Chưa Active",
            marker_color="#ddd",
            marker_line=dict(color="#bbb", width=1),
            text=[f"{int(v):,}" for v in (bu_df["total"] - bu_df["active"])],
            textposition="inside",
            textfont=dict(size=9, color="#666"),
            hovertemplate="<b>%{y}</b><br>Chưa Active: %{x:,}<extra></extra>",
        ))
        
        # Add % label at end
        for _, r in bu_df.iterrows():
            fig_bu.add_annotation(
                x=r["total"] + 50,
                y=r["bu_name"],
                text=f"{r['pct']:.1f}%",
                showarrow=False,
                font=dict(size=10, color="#333", family="Inter"),
            )
        
        fig_bu.update_layout(
            **ibcs_layout(
                height=250,
                margin=dict(l=120, r=80, t=10, b=30),
                barmode='stack',
                showlegend=True,
            ),
            legend=dict(
                orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                font=dict(size=10),
            ),
        )
        fig_bu.update_xaxes(showgrid=True, gridcolor="#eee")
        fig_bu.update_yaxes(showgrid=False)
        
        st.plotly_chart(fig_bu, use_container_width=True)
    
    st.markdown('</div>', unsafe_allow_html=True)


with col_cum:
    st.markdown('<div class="ibcs-section">', unsafe_allow_html=True)
    st.markdown("<h3>TĂNG TRƯỞNG TÍCH LŨY (CUMULATIVE)</h3>", unsafe_allow_html=True)
    
    # Cumulative absolute active users over time
    cum_data = []
    for d in all_dates:
        active_set_d = active_by_date[d]
        total_in_filter = len(df_filtered)
        active_in_filter = int(df_filtered['employee_id'].isin(active_set_d).sum())
        cum_data.append({"date": d, "active": active_in_filter, "total": total_in_filter})
    
    df_cum = pd.DataFrame(cum_data)
    
    fig_cum = go.Figure()
    
    # Area chart for cumulative active (IBCS style – filled gray)
    fig_cum.add_trace(go.Scatter(
        x=df_cum["date"], y=df_cum["active"],
        fill='tozeroy',
        fillcolor='rgba(26,26,26,0.08)',
        line=dict(color=IBCS["AC"], width=2),
        mode='lines+markers',
        marker=dict(size=4, color=IBCS["AC"]),
        text=[f"{v:,}" for v in df_cum["active"]],
        hovertemplate="<b>%{x}</b><br>Active: %{text}<extra></extra>",
    ))
    
    # Reference line for total
    if total_in_filter > 0:
        fig_cum.add_hline(
            y=total_in_filter,
            line=dict(color=IBCS["PY"], width=1, dash="dot"),
            annotation_text=f"Tổng HC: {total_in_filter:,}",
            annotation_position="top right",
            annotation_font=dict(size=9, color="#888"),
        )
    
    fig_cum.update_layout(
        **ibcs_layout(
            height=250,
            margin=dict(l=60, r=20, t=10, b=50),
        ),
    )
    fig_cum.update_xaxes(type="category", tickangle=-45, tickfont=dict(size=8))
    fig_cum.update_yaxes(showgrid=True, gridcolor="#eee")
    
    st.plotly_chart(fig_cum, use_container_width=True)
    
    st.markdown('</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 6: DAILY NET NEW USERS – IBCS Bar Chart
# ═══════════════════════════════════════════════════════════════════════════
st.markdown('<div class="ibcs-section">', unsafe_allow_html=True)
st.markdown("<h3>USER MỚI THEO NGÀY (DAILY NET ADDITIONS)</h3>", unsafe_allow_html=True)

if len(df_trend) > 1:
    df_daily = df_trend.iloc[1:].copy()  # skip first day (no previous)
    
    avg_new = df_daily["new_users"].mean()
    max_day = df_daily.loc[df_daily["new_users"].idxmax()]
    
    st.markdown(
        f'<p class="section-msg">Trung bình: <b>{avg_new:,.0f}</b> user mới/ngày · '
        f'Cao nhất: <b>{int(max_day["new_users"]):,}</b> user ({max_day["date"]})</p>',
        unsafe_allow_html=True
    )
    
    fig_daily = go.Figure()
    
    bar_colors = [IBCS["AC"] if v >= avg_new else IBCS["PY"] for v in df_daily["new_users"]]
    
    fig_daily.add_trace(go.Bar(
        x=df_daily["date"],
        y=df_daily["new_users"],
        marker_color=bar_colors,
        text=[f"{int(v):,}" for v in df_daily["new_users"]],
        textposition="outside",
        textfont=dict(size=9),
        hovertemplate="<b>%{x}</b><br>User mới: %{y:,}<extra></extra>",
    ))
    
    # Average reference line
    fig_daily.add_hline(
        y=avg_new,
        line=dict(color=IBCS["ACCENT"], width=1.5, dash="dash"),
        annotation_text=f"Trung bình: {avg_new:,.0f}",
        annotation_position="top right",
        annotation_font=dict(size=9, color=IBCS["ACCENT"]),
    )
    
    fig_daily.update_layout(
        **ibcs_layout(
            height=280,
            margin=dict(l=50, r=20, t=10, b=50),
        ),
        bargap=0.25,
    )
    fig_daily.update_xaxes(type="category", tickangle=-45)
    
    st.plotly_chart(fig_daily, use_container_width=True)

st.markdown('</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
# SECTION 7: ADOPTION DISTRIBUTION HISTOGRAM
# ═══════════════════════════════════════════════════════════════════════════
st.markdown('<div class="ibcs-section">', unsafe_allow_html=True)
st.markdown("<h3>PHÂN BỐ TỶ LỆ ADOPTION THEO ĐƠN VỊ</h3>", unsafe_allow_html=True)

if len(df_filtered) > 0:
    # Use department level for distribution
    dist_col = "department_name_vn"
    df_dist_temp = df_filtered.copy()
    df_dist_temp["is_active"] = df_dist_temp["employee_id"].isin(active_by_date[selected_date])
    
    dist_df = df_dist_temp.groupby(dist_col).agg(
        total=("employee_id", "count"),
        active=("is_active", "sum"),
    ).reset_index()
    dist_df = dist_df[dist_df["total"] >= 2]  # Filter meaningful units
    dist_df["pct"] = (dist_df["active"] / dist_df["total"] * 100).round(1)
    
    if len(dist_df) > 0:
        # Categorize
        bins = [0, 25, 50, 75, 100.1]
        labels = ["0-25%", "26-50%", "51-75%", "76-100%"]
        dist_df["bucket"] = pd.cut(dist_df["pct"], bins=bins, labels=labels, right=False)
        bucket_counts = dist_df["bucket"].value_counts().reindex(labels, fill_value=0)
        
        median_pct = dist_df["pct"].median()
        st.markdown(
            f'<p class="section-msg">Phân bố tỷ lệ adoption của {len(dist_df)} phòng ban · Median: {median_pct:.1f}%</p>',
            unsafe_allow_html=True
        )
        
        col_hist, col_stats = st.columns([3, 1])
        
        with col_hist:
            fig_hist = go.Figure()
            
            hist_colors = ["#b30000", "#cc6600", IBCS["PY"], IBCS["AC"]]
            
            fig_hist.add_trace(go.Bar(
                x=labels,
                y=bucket_counts.values,
                marker_color=hist_colors,
                text=[f"{int(v)}" for v in bucket_counts.values],
                textposition="outside",
                textfont=dict(size=11, color="#333"),
                hovertemplate="<b>%{x}</b><br>Số đơn vị: %{y}<extra></extra>",
            ))
            
            fig_hist.update_layout(
                **ibcs_layout(
                    height=260,
                    margin=dict(l=50, r=20, t=10, b=40),
                ),
                bargap=0.15,
            )
            fig_hist.update_xaxes(title_text="Nhóm tỷ lệ Adoption", title_font=dict(size=10))
            fig_hist.update_yaxes(title_text="Số đơn vị (Phòng Ban)", title_font=dict(size=10), dtick=5)
            
            st.plotly_chart(fig_hist, use_container_width=True)
        
        with col_stats:
            st.markdown(f"""
            <div style="padding:16px 0; font-size:0.82rem; line-height:2.2;">
                <strong style="font-size:0.72rem; text-transform:uppercase; letter-spacing:0.04em; color:#888;">THỐNG KÊ</strong><br/>
                Trung bình: <strong>{dist_df['pct'].mean():.1f}%</strong><br/>
                Median: <strong>{median_pct:.1f}%</strong><br/>
                Cao nhất: <strong>{dist_df['pct'].max():.1f}%</strong><br/>
                Thấp nhất: <strong>{dist_df['pct'].min():.1f}%</strong><br/>
                Std Dev: <strong>{dist_df['pct'].std():.1f}%</strong><br/>
                ≥75%: <strong>{len(dist_df[dist_df['pct'] >= 75])}</strong> đơn vị<br/>
                <50%: <strong>{len(dist_df[dist_df['pct'] < 50])}</strong> đơn vị
            </div>
            """, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
# FOOTER
# ═══════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="report-footer">
<strong>IBCS</strong> · International Business Communication Standards · 
Data source: <code>[Gtalk Expansion] Workforce Analysis</code> · Developed by <b>EX Team</b> · {datetime.now().strftime('%d/%m/%Y %H:%M')}
</div>
""", unsafe_allow_html=True)