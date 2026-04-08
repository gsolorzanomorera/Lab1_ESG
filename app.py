import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ─── Page config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Environmental Emissions Report",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── McKinsey-style CSS ──────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700&family=IBM+Plex+Sans:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

  [data-testid="stAppViewContainer"] {
    background: #FAFAF8;
    color: #1a1a2e;
    font-family: 'IBM Plex Sans', sans-serif;
  }
  [data-testid="stSidebar"] {
    background: #003366;
    border-right: none;
  }
  [data-testid="stSidebar"] * { color: #e8edf5 !important; }
  [data-testid="stSidebar"] input,
  [data-testid="stSidebar"] .stNumberInput input,
  [data-testid="stSidebar"] .stTextInput input {
    background: #004080 !important;
    color: #e8edf5 !important;
    border: 1px solid #336699 !important;
    border-radius: 2px !important;
  }
  [data-testid="stSidebar"] label {
    font-size: 0.78rem !important;
    letter-spacing: 0.06em !important;
    text-transform: uppercase !important;
    font-weight: 500 !important;
    color: #99bbdd !important;
  }
  [data-testid="stSidebar"] hr { border-color: #336699 !important; opacity: 0.4; }

  #MainMenu, footer, header { visibility: hidden; }
  [data-testid="stToolbar"] { display: none; }

  .stTabs [data-baseweb="tab-list"] {
    background: transparent;
    border-bottom: 2px solid #d0d5dd;
    gap: 0;
    padding: 0;
  }
  .stTabs [data-baseweb="tab"] {
    background: transparent;
    border: none;
    border-bottom: 3px solid transparent;
    border-radius: 0;
    color: #667085;
    font-family: 'IBM Plex Sans', sans-serif;
    font-size: 0.82rem;
    font-weight: 500;
    letter-spacing: 0.07em;
    text-transform: uppercase;
    padding: 12px 24px;
    margin-bottom: -2px;
  }
  .stTabs [aria-selected="true"] {
    background: transparent !important;
    border-bottom: 3px solid #003366 !important;
    color: #003366 !important;
    font-weight: 600 !important;
  }
  .stTabs [data-baseweb="tab-panel"] { padding-top: 28px; }

  .stButton button {
    background: #003366;
    color: white;
    border: none;
    border-radius: 2px;
    font-family: 'IBM Plex Sans', sans-serif;
    font-size: 0.8rem;
    font-weight: 500;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    padding: 10px 24px;
  }
  .stButton button:hover { background: #004080; }

  .streamlit-expanderHeader {
    font-size: 0.8rem !important;
    letter-spacing: 0.06em !important;
    text-transform: uppercase !important;
    color: #667085 !important;
    font-weight: 500 !important;
  }
  hr { border: none; border-top: 1px solid #d0d5dd; margin: 28px 0; }
</style>
""", unsafe_allow_html=True)

# ─── Palette & chart defaults ────────────────────────────────────────────────
C = {
    "navy":  "#003366",
    "teal":  "#006B6B",
    "sky":   "#0077B6",
    "slate": "#4A5568",
    "stone": "#667085",
    "rule":  "#d0d5dd",
    "s1":    "#003366",
    "s2":    "#0077B6",
    "s3":    "#006B6B",
    "total": "#2D3748",
    "bau":   "#C53030",
    "nz":    "#006B6B",
    "up":    "#C53030",
    "down":  "#276749",
}

BASE_LAYOUT = dict(
    plot_bgcolor="white",
    paper_bgcolor="white",
    font=dict(family="IBM Plex Sans, sans-serif", color="#2D3748", size=11),
    margin=dict(t=52, b=56, l=64, r=40),
    showlegend=True,
    legend=dict(
        bgcolor="white", bordercolor="#d0d5dd", borderwidth=1,
        font=dict(size=10, color="#4A5568"),
        orientation="h", yanchor="bottom", y=-0.28, xanchor="left", x=0
    ),
    xaxis=dict(showgrid=False, zeroline=False, showline=True, linecolor="#d0d5dd",
               tickfont=dict(size=10, color="#667085"), ticks="outside", ticklen=4),
    yaxis=dict(showgrid=True, gridcolor="#EBEBEB", zeroline=False, showline=False,
               tickfont=dict(size=10, color="#667085")),
    hoverlabel=dict(bgcolor="white", bordercolor="#d0d5dd",
                    font=dict(family="IBM Plex Sans", size=11, color="#2D3748")),
)

# ─── UI helpers ───────────────────────────────────────────────────────────────
def fmt_m(val):
    if val is None or (isinstance(val, float) and np.isnan(val)): return "—"
    if abs(val) >= 1_000_000: return f"{val/1_000_000:.1f}M"
    if abs(val) >= 1_000: return f"{val/1_000:.0f}K"
    return f"{val:,.0f}"

def pct_chg(new, old):
    if not old or old == 0: return None
    return (new - old) / abs(old) * 100

def delta_html(pct):
    if pct is None: return '<span style="color:#667085">—</span>'
    good = pct < 0
    color = C["down"] if good else C["up"]
    sign = "▲" if pct > 0 else "▼"
    return (f'<span style="color:{color};font-weight:600">{sign} {abs(pct):.1f}%</span>'
            f' <span style="color:#667085;font-size:0.8em">vs. prior year</span>')

def insight_header(headline, subtext=""):
    sub = (f'<div style="font-family:IBM Plex Sans;font-size:0.82rem;color:#667085;'
           f'margin-top:4px;font-weight:400">{subtext}</div>') if subtext else ""
    return f"""
    <div style="margin:32px 0 20px;padding-bottom:12px;border-bottom:2px solid #003366;">
      <div style="font-family:Playfair Display,serif;font-size:1.15rem;font-weight:700;
                  color:#003366;line-height:1.3">{headline}</div>
      {sub}
    </div>"""

def kpi_block(label, value, unit, delta_pct=None, note=""):
    delta = delta_html(delta_pct) if delta_pct is not None else ""
    note_html = (f'<div style="font-size:0.75rem;color:#667085;margin-top:2px">{note}</div>'
                 if note else "")
    return f"""
    <div style="border-top:3px solid #003366;padding:16px 0 12px;">
      <div style="font-family:IBM Plex Sans;font-size:0.7rem;font-weight:600;
                  letter-spacing:0.1em;text-transform:uppercase;color:#667085">{label}</div>
      <div style="font-family:Playfair Display,serif;font-size:2.2rem;font-weight:700;
                  color:#003366;line-height:1.1;margin:6px 0 2px">{value}</div>
      <div style="font-family:IBM Plex Mono,monospace;font-size:0.75rem;color:#4A5568">{unit}</div>
      <div style="margin-top:6px;font-size:0.82rem">{delta}</div>
      {note_html}
    </div>"""

def section_rule(text):
    return f"""
    <div style="display:flex;align-items:center;gap:12px;margin:36px 0 20px">
      <div style="width:32px;height:2px;background:#003366;flex-shrink:0"></div>
      <div style="font-family:IBM Plex Sans;font-size:0.72rem;font-weight:600;
                  letter-spacing:0.12em;text-transform:uppercase;color:#003366;
                  white-space:nowrap">{text}</div>
      <div style="flex:1;height:1px;background:#d0d5dd"></div>
    </div>"""

def apply_layout(fig, **kw):
    layout = dict(BASE_LAYOUT)
    layout.update(kw)
    fig.update_layout(**layout)
    return fig

# ─── Excel parser ─────────────────────────────────────────────────────────────
def parse_excel(file):
    wb = pd.read_excel(file, sheet_name=None, header=None)
    data = {}

    # Assumptions
    try:
        df = wb.get("Assumptions", pd.DataFrame()).astype(str)
        for _, row in df.iterrows():
            rs = " ".join(row.fillna("").astype(str))
            if "Company Name" in rs:
                v = row.dropna(); data["company"] = str(v.iloc[1]) if len(v) > 1 else "Unknown"
            if "Industry" in rs and "Sector" in rs:
                v = row.dropna(); data["industry"] = str(v.iloc[1]) if len(v) > 1 else ""
            if "Annual Revenue" in rs:
                v = pd.to_numeric(row, errors="coerce").dropna()
                if len(v): data["revenue"] = float(v.iloc[0])
            if "Number of Employees" in rs:
                v = pd.to_numeric(row, errors="coerce").dropna()
                if len(v): data["employees"] = float(v.iloc[0])
            if "Reporting Year (Current)" in rs:
                v = pd.to_numeric(row, errors="coerce").dropna()
                if len(v): data["current_year"] = int(v.iloc[0])
    except Exception: pass

    # Summary
    try:
        df = wb.get("Summary & Intensity", pd.DataFrame()).astype(str)
        for _, row in df.iterrows():
            rs = " ".join(row.fillna("").astype(str))
            nums = pd.to_numeric(row, errors="coerce").dropna()
            if "Scope 1" in rs and "Direct" in rs and len(nums):
                data["s1_current"] = float(nums.iloc[0])
            if "Scope 2" in rs and "Market" in rs and "PRIMARY" in rs and len(nums):
                data["s2_mb_current"] = float(nums.iloc[0])
            if "Scope 2" in rs and "Location" in rs and "SECONDARY" in rs and len(nums):
                data["s2_lb_current"] = float(nums.iloc[0])
            if "Scope 3" in rs and "Value Chain" in rs and len(nums):
                data["s3_current"] = float(nums.iloc[0])
    except Exception: pass

    # Trajectory historical
    try:
        df = wb.get("Reduction Trajectory", pd.DataFrame())
        raw = df.astype(str)
        yr_r = s1r = s2lbr = s2mbr = s3r = totr = None
        for i, row in raw.iterrows():
            rs = " ".join(row.fillna("").astype(str))
            if "Scope / Year" in rs: yr_r = i
            if "Scope 1" in rs and "Direct" in rs: s1r = i
            if "location based" in rs.lower() and "Scope 2" in rs: s2lbr = i
            if "market based" in rs.lower() and "Scope 2" in rs: s2mbr = i
            if "Scope 3" in rs and "Value Chain" in rs: s3r = i
            if "TOTAL" in rs and "S1+S2+S3" in rs: totr = i

        if yr_r is not None:
            def extract(ridx):
                if ridx is None: return {}
                vals = pd.to_numeric(df.iloc[ridx], errors="coerce")
                yv   = pd.to_numeric(df.iloc[yr_r], errors="coerce")
                return {int(yv.iloc[c]): float(vals.iloc[c])
                        for c in range(len(vals))
                        if not pd.isna(yv.iloc[c]) and not pd.isna(vals.iloc[c])
                        and 2015 <= yv.iloc[c] <= 2035}
            data["hist_s1"]    = extract(s1r)
            data["hist_s2lb"]  = extract(s2lbr)
            data["hist_s2mb"]  = extract(s2mbr)
            data["hist_s3"]    = extract(s3r)
            data["hist_total"] = extract(totr)
            yv = pd.to_numeric(df.iloc[yr_r], errors="coerce").dropna()
            data["hist_years"] = sorted([int(y) for y in yv if 2015 <= y <= 2035])
    except Exception: pass

    # Trajectories
    try:
        df = wb.get("Reduction Trajectory", pd.DataFrame())
        raw = df.astype(str)
        bau_r = nz_r = sbti_r = s3t_r = ty_r = None
        for i, row in raw.iterrows():
            rs = " ".join(row.fillna("").astype(str))
            if "Business as usual" in rs: bau_r = i
            if "Net Zero Commitment" in rs and bau_r and not nz_r: nz_r = i
            if "Near-Term Science-Based" in rs: sbti_r = i
            if "Scope 3 Reduction Target" in rs: s3t_r = i
            if "Category / Year" in rs: ty_r = i

        if ty_r is not None:
            def ext_t(ridx):
                if ridx is None: return {}
                vals = pd.to_numeric(df.iloc[ridx], errors="coerce")
                yv   = pd.to_numeric(df.iloc[ty_r], errors="coerce")
                return {int(yv.iloc[c]): float(vals.iloc[c])
                        for c in range(len(vals))
                        if not pd.isna(yv.iloc[c]) and not pd.isna(vals.iloc[c])
                        and 2020 <= yv.iloc[c] <= 2035}
            data["traj_bau"]  = ext_t(bau_r)
            data["traj_nz"]   = ext_t(nz_r)
            data["traj_sbti"] = ext_t(sbti_r)
            data["traj_s3"]   = ext_t(s3t_r)
    except Exception: pass

    return data

# ─── Default data ─────────────────────────────────────────────────────────────
DEFAULT = {
    "company": "Your Company", "industry": "Technology",
    "revenue": 245100, "employees": 228000, "current_year": 2024,
    "s1_current": 143510, "s2_mb_current": 259090,
    "s2_lb_current": 9955368, "s3_current": 15140000,
    "hist_years": [2020, 2021, 2022, 2023, 2024],
    "hist_s1":   {2020:118100,   2021:123704,  2022:139413,  2023:144960,  2024:143510},
    "hist_s2mb": {2020:456119,   2021:429405,  2022:288029,  2023:393134,  2024:259090},
    "hist_s2lb": {2020:4328916,  2021:5010667, 2022:6381250, 2023:8077403, 2024:9955368},
    "hist_s3":   {2020:11796000, 2021:13576000,2022:15916000,2023:16397000,2024:15140000},
    "hist_total":{2020:12370000, 2021:14129000,2022:16343000,2023:16935000,2024:15543000},
    "traj_bau":  {2024:12370000,2025:16456070,2026:17422778,2027:18446275,2028:19529898,2029:20677178,2030:21891855},
    "traj_nz":   {2024:12370000,2025:11906125,2026:11442250,2027:10978375,2028:10514500,2029:10050625,2030:9586750},
    "traj_sbti": {2024:15543000,2025:12952500,2026:10362000,2027:7771500, 2028:5181000, 2029:2590500, 2030:0},
    "traj_s3":   {2024:11796000,2025:11206200,2026:10616400,2027:10026600,2028:9436800, 2029:8847000, 2030:8257200},
}

# ─── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="padding:20px 0 8px">
      <div style="font-family:Playfair Display,serif;font-size:1.2rem;font-weight:700;
                  color:white;line-height:1.2">GHG Emissions<br>Intelligence</div>
      <div style="font-size:0.7rem;letter-spacing:0.1em;text-transform:uppercase;
                  color:#99bbdd;margin-top:4px">GHG Protocol Framework</div>
    </div>""", unsafe_allow_html=True)
    st.divider()

    mode = st.radio("Data Source", ["📤 Upload Excel", "✏️ Manual Entry"])
    d = dict(DEFAULT)

    if mode == "📤 Upload Excel":
        uploaded = st.file_uploader("Lab1_dashboard.xlsx", type=["xlsx"])
        if uploaded:
            file_id = f"{uploaded.name}_{uploaded.size}"
            if st.session_state.get("_file_id") != file_id:
                with st.spinner("Parsing…"):
                    parsed = parse_excel(uploaded)
                    st.session_state["_parsed"] = parsed
                    st.session_state["_file_id"] = file_id
            d.update({k: v for k, v in st.session_state["_parsed"].items() if v})
            st.success(f"✅ Loaded: {d.get(chr(39)+'company'+chr(39), 'file')}")
        elif "_parsed" in st.session_state:
            d.update({k: v for k, v in st.session_state["_parsed"].items() if v})
            st.caption(f"Using: {d.get(chr(39)+'company'+chr(39), 'uploaded file')}")
        else:
            st.caption("Demo data: Microsoft 2020–2024")
    else:
        st.markdown("**Company**")
        d["company"]       = st.text_input("Name", d["company"])
        d["industry"]      = st.text_input("Industry", d["industry"])
        d["current_year"]  = st.number_input("Reporting Year", value=int(d["current_year"]), step=1)
        d["revenue"]       = st.number_input("Revenue ($M)", value=float(d["revenue"]), step=1000.0)
        d["employees"]     = st.number_input("Employees (FTE)", value=float(d["employees"]), step=1000.0)
        st.markdown("**Current Year Emissions (tCO₂e)**")
        d["s1_current"]    = st.number_input("Scope 1", value=float(d["s1_current"]), step=1000.0)
        d["s2_mb_current"] = st.number_input("Scope 2 (Market)", value=float(d["s2_mb_current"]), step=1000.0)
        d["s3_current"]    = st.number_input("Scope 3", value=float(d["s3_current"]), step=100000.0)

    st.divider()
    st.markdown("**Targets**")
    nz_year   = st.number_input("Net Zero Year", value=2030, step=1)
    nz_target = st.number_input("Net Zero Target (tCO₂e)", value=8659000, step=100000)

# ─── Derived values ───────────────────────────────────────────────────────────
total_cy = d["s1_current"] + d["s2_mb_current"] + d["s3_current"]
cy = d["current_year"]; py = cy - 1
hist_yrs5 = sorted([y for y in d["hist_years"] if y <= cy])[-5:]

d1_s1  = pct_chg(d["s1_current"],    d["hist_s1"].get(py))
d1_s2  = pct_chg(d["s2_mb_current"], d["hist_s2mb"].get(py))
d1_s3  = pct_chg(d["s3_current"],    d["hist_s3"].get(py))
d1_tot = pct_chg(total_cy,           d["hist_total"].get(py))

int_rev = total_cy / d["revenue"]   if d["revenue"]   else 0
int_emp = total_cy / d["employees"] if d["employees"] else 0
s3_pct  = d["s3_current"] / total_cy * 100 if total_cy else 0
s1_pct  = d["s1_current"] / total_cy * 100 if total_cy else 0
s2_pct  = d["s2_mb_current"] / total_cy * 100 if total_cy else 0

# ─── Page header ─────────────────────────────────────────────────────────────
st.markdown(f"""
<div style="border-bottom:3px solid #003366;padding:32px 0 24px;margin-bottom:0">
  <div style="display:flex;justify-content:space-between;align-items:flex-end">
    <div>
      <div style="font-family:IBM Plex Sans;font-size:0.72rem;font-weight:600;
                  letter-spacing:0.14em;text-transform:uppercase;color:#667085;
                  margin-bottom:8px">Carbon Footprint Assessment · {cy}</div>
      <div style="font-family:Playfair Display,serif;font-size:2.4rem;font-weight:700;
                  color:#003366;line-height:1.1">{d['company']}</div>
      <div style="font-family:IBM Plex Sans;font-size:0.9rem;color:#4A5568;margin-top:8px">
        {d['industry']} &nbsp;·&nbsp; Revenue: <b>${d['revenue']:,.0f}M</b>
        &nbsp;·&nbsp; Employees: <b>{int(d['employees']):,}</b>
      </div>
    </div>
    <div style="text-align:right">
      <div style="font-family:IBM Plex Mono,monospace;font-size:0.7rem;
                  color:#667085;letter-spacing:0.06em">GHG PROTOCOL FRAMEWORK<br>SCOPE 1 · 2 · 3</div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# ─── KPI row ─────────────────────────────────────────────────────────────────
st.markdown(insight_header(
    f"Scope 3 dominates — {s3_pct:.0f}% of total footprint driven by value chain",
    f"Total inventory: {fmt_m(total_cy)} tCO₂e  ·  Reporting year {cy}  ·  Market-based Scope 2"
), unsafe_allow_html=True)

k1, k2, k3, k4, k5 = st.columns(5)
k1.markdown(kpi_block("Scope 1 · Direct",       fmt_m(d["s1_current"]),    "tCO₂e", d1_s1),  unsafe_allow_html=True)
k2.markdown(kpi_block("Scope 2 · Market-Based", fmt_m(d["s2_mb_current"]), "tCO₂e", d1_s2),  unsafe_allow_html=True)
k3.markdown(kpi_block("Scope 3 · Value Chain",  fmt_m(d["s3_current"]),    "tCO₂e", d1_s3),  unsafe_allow_html=True)
k4.markdown(kpi_block("Total S1 + S2 + S3",     fmt_m(total_cy),           "tCO₂e", d1_tot), unsafe_allow_html=True)
k5.markdown(kpi_block("Carbon Intensity",        f"{int_rev:.1f}", "tCO₂e / $M revenue",
                       note=f"{int_emp:.1f} tCO₂e / employee"), unsafe_allow_html=True)

st.markdown("<hr>", unsafe_allow_html=True)

# ─── Tabs ────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "Emission Trend", "Reduction Trajectory", "Carbon Intensity", "Scope Breakdown"
])

# ── Tab 1: Emission Trend ─────────────────────────────────────────────────────
with tab1:
    s1v = [d["hist_s1"].get(y)    for y in hist_yrs5]
    s2v = [d["hist_s2mb"].get(y)  for y in hist_yrs5]
    s3v = [d["hist_s3"].get(y)    for y in hist_yrs5]
    tv  = [d["hist_total"].get(y) for y in hist_yrs5]

    st.markdown(insight_header(
        "Absolute emissions declined in 2024 after peaking in 2023",
        "Scope 3 value chain emissions are the primary lever — representing over 97% of inventory"
    ), unsafe_allow_html=True)

    col_bar, col_yoy = st.columns([3, 2])

    with col_bar:
        fig = go.Figure()
        bw = 0.22
        fig.add_trace(go.Bar(name="Scope 1", x=[y-bw for y in hist_yrs5], y=s1v,
                             marker_color=C["s1"], width=bw,
                             text=[fmt_m(v) for v in s1v], textposition="outside",
                             textfont=dict(size=9, color=C["navy"])))
        fig.add_trace(go.Bar(name="Scope 2 (Market)", x=hist_yrs5, y=s2v,
                             marker_color=C["s2"], width=bw,
                             text=[fmt_m(v) for v in s2v], textposition="outside",
                             textfont=dict(size=9, color=C["sky"])))
        fig.add_trace(go.Bar(name="Scope 3", x=[y+bw for y in hist_yrs5], y=s3v,
                             marker_color=C["s3"], width=bw,
                             text=[fmt_m(v) for v in s3v], textposition="outside",
                             textfont=dict(size=9, color=C["teal"])))
        fig.add_trace(go.Scatter(name="Total", x=hist_yrs5, y=tv,
                                 mode="lines+markers+text",
                                 text=[fmt_m(v) for v in tv], textposition="top center",
                                 textfont=dict(size=9, color=C["total"], family="IBM Plex Mono"),
                                 line=dict(color=C["total"], width=2, dash="dot"),
                                 marker=dict(size=7, color=C["total"],
                                             line=dict(color="white", width=2))))
        apply_layout(fig, title="Absolute Emissions by Scope (tCO₂e)", height=420,
                     xaxis=dict(tickvals=hist_yrs5, tickformat="d", showgrid=False,
                                zeroline=False, showline=True, linecolor="#d0d5dd",
                                tickfont=dict(size=10, color="#667085")),
                     yaxis=dict(showgrid=True, gridcolor="#EBEBEB", zeroline=False,
                                tickfont=dict(size=10, color="#667085"), tickformat=".2s"))
        st.plotly_chart(fig, use_container_width=True)

    with col_yoy:
        yoy_s1  = [None]+[pct_chg(s1v[i], s1v[i-1])  for i in range(1, len(s1v))]
        yoy_s3  = [None]+[pct_chg(s3v[i], s3v[i-1])  for i in range(1, len(s3v))]
        yoy_tot = [None]+[pct_chg(tv[i],  tv[i-1])    for i in range(1, len(tv))]

        fig2 = go.Figure()
        fig2.add_hline(y=0, line_color="#d0d5dd", line_width=1.5)
        fig2.add_trace(go.Scatter(name="Scope 1 YoY", x=hist_yrs5, y=yoy_s1,
                                  mode="lines+markers", line=dict(color=C["s1"], width=1.5, dash="dot"),
                                  marker=dict(size=5)))
        fig2.add_trace(go.Scatter(name="Scope 3 YoY", x=hist_yrs5, y=yoy_s3,
                                  mode="lines+markers", line=dict(color=C["s3"], width=1.5, dash="dot"),
                                  marker=dict(size=5)))
        fig2.add_trace(go.Scatter(name="Total YoY", x=hist_yrs5, y=yoy_tot,
                                  mode="lines+markers+text",
                                  text=[f"{v:.1f}%" if v else "" for v in yoy_tot],
                                  textposition="top center",
                                  textfont=dict(size=9, color=C["total"]),
                                  line=dict(color=C["total"], width=2),
                                  marker=dict(size=7, color=C["total"],
                                              line=dict(color="white", width=2))))
        apply_layout(fig2, title="Year-on-Year Change (%)", height=420,
                     xaxis=dict(tickvals=hist_yrs5, tickformat="d", showgrid=False,
                                zeroline=False, showline=True, linecolor="#d0d5dd",
                                tickfont=dict(size=10, color="#667085")),
                     yaxis=dict(showgrid=True, gridcolor="#EBEBEB", zeroline=False,
                                tickfont=dict(size=10, color="#667085"), ticksuffix="%"))
        st.plotly_chart(fig2, use_container_width=True)

    with st.expander("View Data Table"):
        tbl = pd.DataFrame({"Year": hist_yrs5, "Scope 1": s1v,
                             "Scope 2 MB": s2v, "Scope 3": s3v, "Total": tv})
        st.dataframe(tbl.style.format(
            {c: "{:,.0f}" for c in tbl.columns if c != "Year"}),
            use_container_width=True, hide_index=True)

# ── Tab 2: Reduction Trajectory ────────────────────────────────────────────────
with tab2:
    bau_val = d.get("traj_bau", {}).get(nz_year, 0)
    gap = (bau_val - nz_target) / 1e6
    st.markdown(insight_header(
        f"A {gap:.0f}M tCO₂e gap separates business-as-usual from the {nz_year} Net Zero target",
        "Accelerated decarbonisation across all scopes is required to meet SBTi-aligned pathways"
    ), unsafe_allow_html=True)

    fig3 = go.Figure()
    fig3.add_trace(go.Scatter(
        name="Historical Actual", x=hist_yrs5,
        y=[d["hist_total"].get(y) for y in hist_yrs5],
        mode="lines+markers", line=dict(color=C["navy"], width=3),
        marker=dict(size=8, color=C["navy"], line=dict(color="white", width=2)),
    ))

    def add_line(name, key, color, dash, width=1.5):
        td = d.get(key, {})
        if not td: return
        yrs = sorted(td.keys())
        fig3.add_trace(go.Scatter(
            name=name, x=yrs, y=[td[y] for y in yrs],
            mode="lines", line=dict(color=color, width=width, dash=dash),
        ))

    add_line("Business As Usual",    "traj_bau",  C["bau"], "dash")
    add_line("Net Zero Pathway",     "traj_nz",   C["nz"],  "solid", 2)
    add_line("SBTi Scope 1+2",       "traj_sbti", C["sky"], "dot")
    add_line("SBTi Scope 3 Target",  "traj_s3",   "#5B5EA6","dot")

    fig3.add_annotation(
        x=nz_year, y=nz_target,
        text=f"Net Zero Target<br><b>{fmt_m(nz_target)} tCO₂e</b>",
        showarrow=True, arrowhead=2, arrowcolor=C["nz"], arrowwidth=1.5,
        font=dict(color=C["nz"], size=10, family="IBM Plex Sans"),
        bgcolor="white", bordercolor=C["nz"], borderwidth=1, borderpad=6,
        ax=80, ay=-60
    )
    apply_layout(fig3, title="Emission Trajectory vs. Reduction Pathways (tCO₂e)", height=460,
                 xaxis=dict(tickformat="d", showgrid=False, zeroline=False,
                            showline=True, linecolor="#d0d5dd",
                            tickfont=dict(size=10, color="#667085")),
                 yaxis=dict(showgrid=True, gridcolor="#EBEBEB", zeroline=False,
                            tickfont=dict(size=10, color="#667085"), tickformat=".2s"))
    st.plotly_chart(fig3, use_container_width=True)

    st.markdown(section_rule("Reduction Targets"), unsafe_allow_html=True)
    st.dataframe(pd.DataFrame([
        {"Target": "Net Zero Commitment",      "Year": nz_year, "Baseline (tCO₂e)": "12,370,000",
         "Reduction Required": "30%", "Target (tCO₂e)": f"{nz_target:,.0f}", "Framework": "GHG Protocol"},
        {"Target": "SBTi Near-Term (S1+S2)",  "Year": 2030,    "Baseline (tCO₂e)": "15,543,000",
         "Reduction Required": "100%","Target (tCO₂e)": "0",              "Framework": "SBTi 1.5°C"},
        {"Target": "SBTi Scope 3 Reduction",  "Year": 2030,    "Baseline (tCO₂e)": "11,796,000",
         "Reduction Required": "50%", "Target (tCO₂e)": "5,898,000",      "Framework": "SBTi value chain"},
    ]), use_container_width=True, hide_index=True)

# ── Tab 3: Carbon Intensity ───────────────────────────────────────────────────
with tab3:
    int_rev_prev = d["hist_total"].get(py, 0) / d["revenue"] if d["revenue"] else 0
    int_chg = pct_chg(int_rev, int_rev_prev)
    dir_word = "improved" if int_chg and int_chg < 0 else "increased"
    st.markdown(insight_header(
        (f"Carbon intensity {dir_word} {abs(int_chg):.1f}% — "
         f"now {int_rev:.1f} tCO₂e per $M revenue") if int_chg else "Carbon intensity ratios",
        "Intensity metrics decouple emissions performance from business growth"
    ), unsafe_allow_html=True)

    int_rev_vals = [d["hist_total"].get(y, 0) / d["revenue"]   if d["revenue"]   else None for y in hist_yrs5]
    int_emp_vals = [d["hist_total"].get(y, 0) / d["employees"] if d["employees"] else None for y in hist_yrs5]

    ci1, ci2 = st.columns(2)

    def intensity_chart(yvals, title, color):
        f = go.Figure()
        f.add_trace(go.Bar(x=hist_yrs5, y=yvals, marker_color=color, opacity=0.85, width=0.5,
                           text=[f"{v:.1f}" if v else "" for v in yvals],
                           textposition="outside",
                           textfont=dict(size=10, color=color, family="IBM Plex Mono")))
        apply_layout(f, title=title, height=360, showlegend=False,
                     margin=dict(t=52, b=40, l=64, r=20),
                     xaxis=dict(tickvals=hist_yrs5, tickformat="d", showgrid=False,
                                zeroline=False, showline=True, linecolor="#d0d5dd",
                                tickfont=dict(size=10, color="#667085")),
                     yaxis=dict(showgrid=True, gridcolor="#EBEBEB", zeroline=False,
                                tickfont=dict(size=10, color="#667085")))
        return f

    with ci1:
        st.plotly_chart(intensity_chart(int_rev_vals,
            "Carbon Intensity — Revenue (tCO₂e / $M)", C["navy"]), use_container_width=True)
    with ci2:
        st.plotly_chart(intensity_chart(int_emp_vals,
            "Carbon Intensity — Employee (tCO₂e / FTE)", C["teal"]), use_container_width=True)

    st.markdown(section_rule("Scope 2: Location-Based vs. Market-Based"), unsafe_allow_html=True)
    st.markdown("""<div style="font-size:0.82rem;color:#667085;margin-bottom:16px;max-width:680px">
    GHG Protocol requires dual reporting. The gap between location-based and market-based figures
    reflects the impact of renewable energy certificates (RECs) and power purchase agreements (PPAs).
    </div>""", unsafe_allow_html=True)

    s2lb_v = [d["hist_s2lb"].get(y) for y in hist_yrs5]
    s2mb_v = [d["hist_s2mb"].get(y) for y in hist_yrs5]
    fig_s2 = go.Figure()
    fig_s2.add_trace(go.Bar(name="Location-Based (grid average)",   x=[y-0.18 for y in hist_yrs5],
                            y=s2lb_v, marker_color=C["slate"], opacity=0.5, width=0.3,
                            text=[fmt_m(v) for v in s2lb_v], textposition="outside",
                            textfont=dict(size=9, color=C["slate"])))
    fig_s2.add_trace(go.Bar(name="Market-Based (REC/PPA adjusted)", x=[y+0.18 for y in hist_yrs5],
                            y=s2mb_v, marker_color=C["sky"], width=0.3,
                            text=[fmt_m(v) for v in s2mb_v], textposition="outside",
                            textfont=dict(size=9, color=C["sky"])))
    apply_layout(fig_s2, title="Scope 2 Dual Reporting (tCO₂e)", height=340,
                 xaxis=dict(tickvals=hist_yrs5, tickformat="d", showgrid=False,
                            zeroline=False, showline=True, linecolor="#d0d5dd",
                            tickfont=dict(size=10, color="#667085")),
                 yaxis=dict(showgrid=True, gridcolor="#EBEBEB", zeroline=False,
                            tickfont=dict(size=10, color="#667085"), tickformat=".2s"))
    st.plotly_chart(fig_s2, use_container_width=True)

# ── Tab 4: Scope Breakdown ────────────────────────────────────────────────────
with tab4:
    st.markdown(insight_header(
        f"Value chain (Scope 3) is the critical focus — {s3_pct:.0f}% of total footprint",
        f"Scope 1: {s1_pct:.1f}%  ·  Scope 2 (market-based): {s2_pct:.1f}%  ·  Scope 3: {s3_pct:.1f}%"
    ), unsafe_allow_html=True)

    col_pie, col_area = st.columns(2)

    with col_pie:
        fig_pie = go.Figure(go.Pie(
            labels=["Scope 1 — Direct", "Scope 2 — Market", "Scope 3 — Value Chain"],
            values=[d["s1_current"], d["s2_mb_current"], d["s3_current"]],
            hole=0.62,
            marker=dict(colors=[C["s1"], C["s2"], C["s3"]],
                        line=dict(color="white", width=3)),
            textinfo="label+percent",
            textfont=dict(family="IBM Plex Sans", size=11, color="#2D3748"),
            direction="clockwise", sort=False,
        ))
        fig_pie.add_annotation(
            text=(f"<b>{fmt_m(total_cy)}</b><br>"
                  f"<span style='font-size:11px;color:#667085'>tCO₂e</span>"),
            x=0.5, y=0.5, showarrow=False,
            font=dict(color="#003366", size=14, family="Playfair Display")
        )
        fig_pie.update_layout(
            title=f"Scope Mix — {cy}", height=420,
            paper_bgcolor="white", plot_bgcolor="white",
            font=dict(family="IBM Plex Sans", color="#2D3748"),
            legend=dict(bgcolor="white", bordercolor="#d0d5dd", borderwidth=1,
                        font=dict(size=10), orientation="h",
                        yanchor="top", y=-0.05),
            margin=dict(t=52, b=40, l=20, r=20)
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    with col_area:
        fig_area = go.Figure()
        fig_area.add_trace(go.Scatter(
            name="Scope 3", x=hist_yrs5,
            y=[d["hist_s3"].get(y, 0) for y in hist_yrs5],
            stackgroup="one", fillcolor="rgba(0,107,107,0.75)",
            line=dict(color=C["s3"], width=0)
        ))
        fig_area.add_trace(go.Scatter(
            name="Scope 2 (Market)", x=hist_yrs5,
            y=[d["hist_s2mb"].get(y, 0) for y in hist_yrs5],
            stackgroup="one", fillcolor="rgba(0,119,182,0.75)",
            line=dict(color=C["s2"], width=0)
        ))
        fig_area.add_trace(go.Scatter(
            name="Scope 1", x=hist_yrs5,
            y=[d["hist_s1"].get(y, 0) for y in hist_yrs5],
            stackgroup="one", fillcolor="rgba(0,51,102,0.85)",
            line=dict(color=C["s1"], width=0)
        ))
        apply_layout(fig_area, title="Stacked Composition Over Time (tCO₂e)", height=420,
                     xaxis=dict(tickvals=hist_yrs5, tickformat="d", showgrid=False,
                                zeroline=False, showline=True, linecolor="#d0d5dd",
                                tickfont=dict(size=10, color="#667085")),
                     yaxis=dict(showgrid=True, gridcolor="#EBEBEB", zeroline=False,
                                tickfont=dict(size=10, color="#667085"), tickformat=".2s"))
        st.plotly_chart(fig_area, use_container_width=True)

# ─── Export ───────────────────────────────────────────────────────────────────
st.markdown("<hr>", unsafe_allow_html=True)
st.markdown(section_rule("Export"), unsafe_allow_html=True)

if st.button("Generate CSV Summary"):
    rows = (
        [["CARBON FOOTPRINT SUMMARY REPORT", "", ""],
         [f"Company: {d['company']}", f"Year: {cy}", f"Industry: {d['industry']}"],
         ["", "", ""],
         ["SCOPE TOTALS", "tCO₂e", ""],
         ["Scope 1 — Direct", d["s1_current"], ""],
         ["Scope 2 — Market-Based", d["s2_mb_current"], ""],
         ["Scope 2 — Location-Based", d["s2_lb_current"], ""],
         ["Scope 3 — Value Chain", d["s3_current"], ""],
         ["Total (S1+S2MB+S3)", total_cy, ""],
         ["", "", ""],
         ["INTENSITY", "", ""],
         ["Carbon Intensity — Revenue", round(int_rev, 2), "tCO₂e / $M"],
         ["Carbon Intensity — Employee", round(int_emp, 2), "tCO₂e / FTE"],
         ["", "", ""],
         ["HISTORICAL TREND", "", ""]] +
        [[y, d["hist_total"].get(y, ""), "tCO₂e"] for y in hist_yrs5]
    )
    csv = pd.DataFrame(rows).to_csv(index=False, header=False)
    st.download_button("⬇ Download CSV", csv,
                       file_name=f"GHG_Report_{d['company']}_{cy}.csv",
                       mime="text/csv")

st.markdown(f"""
<div style="margin-top:40px;padding-top:16px;border-top:1px solid #d0d5dd;
            display:flex;justify-content:space-between;align-items:center">
  <div style="font-family:IBM Plex Sans;font-size:0.72rem;color:#9BA3AF;letter-spacing:0.04em">
    GHG Protocol Corporate Standard &nbsp;·&nbsp;
    Emission factors: EPA CCCL 2023, DEFRA 2023, IEA 2023
  </div>
  <div style="font-family:IBM Plex Mono;font-size:0.7rem;color:#9BA3AF">{cy} REPORT</div>
</div>
""", unsafe_allow_html=True)
