## ============================================================
## GHG Emissions Dashboard — McKinsey-style Streamlit App
## GHG Protocol Framework: Scope 1, 2, 3 reporting
## Supports Excel upload or full manual data entry
## ============================================================

## Import required libraries
import streamlit as st             ## Streamlit: builds the web UI, handles widgets and layout
import pandas as pd                ## Pandas: reads Excel files and manipulates tabular data
import numpy as np                 ## NumPy: used for math helpers like isnan()
import plotly.graph_objects as go  ## Plotly: builds all interactive charts (bars, lines, pies)
from plotly.subplots import make_subplots  ## Plotly helper: multi-panel layouts (imported for extensibility)

## ─── PAGE CONFIGURATION ──────────────────────────────────────────────────────
## Must be the FIRST Streamlit call — sets browser tab title, icon, and layout
st.set_page_config(
    page_title="GHG Emissions Report",  ## Text shown in the browser tab
    page_icon="🌿",                      ## Emoji shown in the browser tab
    layout="wide",                       ## Use full browser width instead of a narrow centered column
    initial_sidebar_state="expanded",    ## Sidebar is open by default when the page loads
)

## ─── CUSTOM CSS STYLING ──────────────────────────────────────────────────────
## Injects CSS into the page to apply McKinsey-style design:
## white background, navy sidebar, clean serif typography, minimal borders
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700&family=IBM+Plex+Sans:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

  [data-testid="stAppViewContainer"] {
    background: #FAFAF8;               /* Warm off-white page background */
    color: #1a1a2e;                    /* Near-black default text color */
    font-family: 'IBM Plex Sans', sans-serif;  /* Default body font */
  }
  [data-testid="stSidebar"] {
    background: #003366;               /* Deep navy sidebar background */
    border-right: none;                /* Remove default border between sidebar and content */
  }
  [data-testid="stSidebar"] * { color: #e8edf5 !important; }  /* Force all sidebar text to light blue-white */
  [data-testid="stSidebar"] input,
  [data-testid="stSidebar"] .stNumberInput input,
  [data-testid="stSidebar"] .stTextInput input {
    background: #004080 !important;   /* Darker navy fill for input fields */
    color: #e8edf5 !important;        /* Light text in inputs */
    border: 1px solid #336699 !important;  /* Subtle blue border */
    border-radius: 2px !important;    /* Minimal rounding — sharp professional look */
  }
  [data-testid="stSidebar"] label {
    font-size: 0.78rem !important;
    letter-spacing: 0.06em !important;
    text-transform: uppercase !important;  /* Small-caps label style */
    font-weight: 500 !important;
    color: #99bbdd !important;         /* Muted blue label color on dark sidebar */
  }
  [data-testid="stSidebar"] hr { border-color: #336699 !important; opacity: 0.4; }  /* Subtle sidebar dividers */

  #MainMenu, footer, header { visibility: hidden; }  /* Hide Streamlit's default menu, footer, and header */
  [data-testid="stToolbar"] { display: none; }       /* Hide the top-right toolbar */

  .stTabs [data-baseweb="tab-list"] {
    background: transparent;
    border-bottom: 2px solid #d0d5dd;  /* Light gray underline running full tab width */
    gap: 0;
    padding: 0;
  }
  .stTabs [data-baseweb="tab"] {
    background: transparent;
    border: none;
    border-bottom: 3px solid transparent;  /* Invisible bottom border reserves space for the active indicator */
    border-radius: 0;
    color: #667085;                    /* Gray text for inactive tabs */
    font-family: 'IBM Plex Sans', sans-serif;
    font-size: 0.82rem;
    font-weight: 500;
    letter-spacing: 0.07em;
    text-transform: uppercase;
    padding: 12px 24px;
    margin-bottom: -2px;               /* Overlap the tab-list border so active tab border replaces it */
  }
  .stTabs [aria-selected="true"] {
    background: transparent !important;
    border-bottom: 3px solid #003366 !important;  /* Navy bottom border marks the active tab */
    color: #003366 !important;
    font-weight: 600 !important;
  }
  .stTabs [data-baseweb="tab-panel"] { padding-top: 28px; }  /* Breathing room between tabs and content */

  .stButton button {
    background: #003366;               /* Navy button fill */
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
  .stButton button:hover { background: #004080; }  /* Slightly lighter navy on hover */

  .streamlit-expanderHeader {
    font-size: 0.8rem !important;
    letter-spacing: 0.06em !important;
    text-transform: uppercase !important;
    color: #667085 !important;
    font-weight: 500 !important;
  }
  hr { border: none; border-top: 1px solid #d0d5dd; margin: 28px 0; }  /* Thin light-gray page dividers */
</style>
""", unsafe_allow_html=True)  ## unsafe_allow_html=True is required to inject raw HTML/CSS

## ─── COLOR PALETTE ───────────────────────────────────────────────────────────
## Central dictionary of all brand colors — change once here, updates everywhere
C = {
    "navy":  "#003366",  ## Primary navy — Scope 1 bars, headers, KPI values
    "teal":  "#006B6B",  ## Deep teal — Scope 3, Net Zero pathway, employee intensity chart
    "sky":   "#0077B6",  ## Sky blue — Scope 2 market-based, SBTi targets
    "slate": "#4A5568",  ## Dark slate — body text, location-based Scope 2 bars
    "stone": "#667085",  ## Medium gray — labels, captions, secondary text
    "rule":  "#d0d5dd",  ## Light gray — dividers, chart gridlines, card borders
    "s1":    "#003366",  ## Scope 1 chart color (alias of navy)
    "s2":    "#0077B6",  ## Scope 2 chart color (alias of sky)
    "s3":    "#006B6B",  ## Scope 3 chart color (alias of teal)
    "total": "#2D3748",  ## Total/combined line — near-black dark slate
    "bau":   "#C53030",  ## Business-as-usual trajectory — muted red (worst case)
    "nz":    "#006B6B",  ## Net Zero pathway — teal (best case)
    "up":    "#C53030",  ## Delta arrow color when emissions increased (red = bad for environment)
    "down":  "#276749",  ## Delta arrow color when emissions decreased (green = good for environment)
}

## ─── BASE CHART LAYOUT ───────────────────────────────────────────────────────
## Shared Plotly layout applied to every chart for visual consistency
## Each chart merges this base with its own overrides (title, height, axis format)
BASE_LAYOUT = dict(
    plot_bgcolor="white",   ## White chart area background
    paper_bgcolor="white",  ## White area outside the axes
    font=dict(family="IBM Plex Sans, sans-serif", color="#2D3748", size=11),  ## Default chart font
    margin=dict(t=52, b=56, l=64, r=40),  ## Chart margins in pixels: top, bottom, left, right
    showlegend=True,        ## Show legend on all charts by default
    legend=dict(
        bgcolor="white", bordercolor="#d0d5dd", borderwidth=1,  ## White legend box with light border
        font=dict(size=10, color="#4A5568"),    ## Small gray legend text
        orientation="h",                        ## Horizontal layout (items side by side)
        yanchor="bottom", y=-0.28,             ## Position legend below the chart
        xanchor="left", x=0                    ## Align to left edge
    ),
    xaxis=dict(
        showgrid=False,      ## No vertical gridlines — cleaner, less visual clutter
        zeroline=False,      ## No zero-line marker on x-axis
        showline=True,       ## Show the x-axis baseline
        linecolor="#d0d5dd", ## Light gray baseline
        tickfont=dict(size=10, color="#667085"),  ## Small gray tick labels
        ticks="outside",     ## Tick marks outside the chart area
        ticklen=4            ## Short tick mark length
    ),
    yaxis=dict(
        showgrid=True,            ## Horizontal gridlines to guide the eye across the chart
        gridcolor="#EBEBEB",      ## Very light gray gridlines (subtle, not distracting)
        zeroline=False,           ## No zero-line
        showline=False,           ## No y-axis border line — gridlines are sufficient
        tickfont=dict(size=10, color="#667085")  ## Small gray tick labels
    ),
    hoverlabel=dict(
        bgcolor="white",          ## White tooltip background
        bordercolor="#d0d5dd",    ## Light border on tooltips
        font=dict(family="IBM Plex Sans", size=11, color="#2D3748")  ## Clean tooltip font
    ),
)

## ─── UTILITY FUNCTIONS ───────────────────────────────────────────────────────

def fmt_m(val):
    ## Format large numbers compactly for display: 15,542,600 → "15.5M", 259,090 → "259K"
    ## Returns an em dash for missing or NaN values
    if val is None or (isinstance(val, float) and np.isnan(val)): return "—"  ## Handle missing data
    if abs(val) >= 1_000_000: return f"{val/1_000_000:.1f}M"  ## Millions with one decimal place
    if abs(val) >= 1_000: return f"{val/1_000:.0f}K"           ## Thousands with no decimal
    return f"{val:,.0f}"                                        ## Small numbers with comma formatting

def pct_chg(new, old):
    ## Calculate percentage change from old to new: ((new - old) / old) × 100
    ## Returns None if old is zero or missing, avoiding division-by-zero errors
    if not old or old == 0: return None        ## Guard against zero or None denominator
    return (new - old) / abs(old) * 100        ## Standard percentage change formula

def delta_html(pct):
    ## Render a colored ▲/▼ arrow badge with a percentage for KPI cards
    ## For emissions, a decrease (▼) is GREEN (good); an increase (▲) is RED (bad)
    if pct is None: return '<span style="color:#667085">—</span>'  ## Show dash when no prior-year data
    good = pct < 0                            ## Emissions falling = environmentally positive
    color = C["down"] if good else C["up"]    ## Green for decrease, red for increase
    sign = "▲" if pct > 0 else "▼"           ## Up arrow for increases, down arrow for decreases
    return (f'<span style="color:{color};font-weight:600">{sign} {abs(pct):.1f}%</span>'
            f' <span style="color:#667085;font-size:0.8em">vs. prior year</span>')  ## Small "vs. prior year" suffix

def insight_header(headline, subtext=""):
    ## Generate a McKinsey-style section heading: bold serif headline with 2px navy bottom border
    ## The navy border visually anchors the headline to the content below it
    ## subtext is optional — a smaller gray explanatory line under the headline
    sub = (f'<div style="font-family:IBM Plex Sans;font-size:0.82rem;color:#667085;'
           f'margin-top:4px;font-weight:400">{subtext}</div>') if subtext else ""  ## Only render if subtext is provided
    return f"""
    <div style="margin:32px 0 20px;padding-bottom:12px;border-bottom:2px solid #003366;">
      <div style="font-family:Playfair Display,serif;font-size:1.15rem;font-weight:700;
                  color:#003366;line-height:1.3">{headline}</div>
      {sub}
    </div>"""

def kpi_block(label, value, unit, delta_pct=None, note=""):
    ## Generate HTML for a KPI metric card in McKinsey consulting style
    ## Structure: navy 3px top accent → small-caps label → large serif value → monospace unit → delta arrow → note
    delta = delta_html(delta_pct) if delta_pct is not None else ""  ## Only show delta badge if percentage was passed
    note_html = (f'<div style="font-size:0.75rem;color:#667085;margin-top:2px">{note}</div>'
                 if note else "")  ## Optional small note (used for the employee intensity sub-figure)
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
    ## Generate a horizontal divider with a short 32px navy bar and an uppercase text label
    ## Used to separate sub-sections within a tab (e.g., the targets table below the trajectory chart)
    return f"""
    <div style="display:flex;align-items:center;gap:12px;margin:36px 0 20px">
      <div style="width:32px;height:2px;background:#003366;flex-shrink:0"></div>
      <div style="font-family:IBM Plex Sans;font-size:0.72rem;font-weight:600;
                  letter-spacing:0.12em;text-transform:uppercase;color:#003366;
                  white-space:nowrap">{text}</div>
      <div style="flex:1;height:1px;background:#d0d5dd"></div>
    </div>"""

def apply_layout(fig, **kw):
    ## Apply BASE_LAYOUT to a Plotly figure, then merge in chart-specific overrides
    ## Using **kw allows callers to override any key in BASE_LAYOUT without rewriting the whole thing
    layout = dict(BASE_LAYOUT)  ## Copy base layout so the original is never mutated
    layout.update(kw)            ## Merge chart-specific keys (title, height, axis format, etc.)
    fig.update_layout(**layout)  ## Apply the merged settings to the Plotly figure
    return fig                   ## Return fig so calls can be chained if needed

## ─── EXCEL PARSER ─────────────────────────────────────────────────────────────
def parse_excel(file):
    ## Parse the Lab1_dashboard.xlsx file and return a dict of structured data
    ## Uses keyword-based row detection instead of fixed row numbers so it works
    ## even if the user adds rows above or rearranges the layout slightly

    wb = pd.read_excel(file, sheet_name=None, header=None)  ## Load ALL sheets; header=None means no automatic header row
    data = {}  ## Accumulate all extracted values here

    ## --- "Assumptions" sheet: company info and financial metrics ---
    try:
        df = wb.get("Assumptions", pd.DataFrame()).astype(str)  ## Get sheet as string DataFrame for keyword searching
        for _, row in df.iterrows():                             ## Loop through every row
            rs = " ".join(row.fillna("").astype(str))           ## Join all cells into one searchable string
            if "Company Name" in rs:
                v = row.dropna()                                 ## Remove empty cells
                data["company"] = str(v.iloc[1]) if len(v) > 1 else "Unknown"  ## Name is always in column index 1
            if "Industry" in rs and "Sector" in rs:
                v = row.dropna()
                data["industry"] = str(v.iloc[1]) if len(v) > 1 else ""  ## Industry is in column index 1
            if "Annual Revenue" in rs:
                v = pd.to_numeric(row, errors="coerce").dropna()  ## Extract only numeric cells
                if len(v): data["revenue"] = float(v.iloc[0])    ## Revenue is the first numeric value on this row
            if "Number of Employees" in rs:
                v = pd.to_numeric(row, errors="coerce").dropna()
                if len(v): data["employees"] = float(v.iloc[0])  ## Headcount is the first numeric value
            if "Reporting Year (Current)" in rs:
                v = pd.to_numeric(row, errors="coerce").dropna()
                if len(v): data["current_year"] = int(v.iloc[0])  ## Reporting year is the first numeric value
    except Exception: pass  ## Silently skip if this sheet is missing or structured differently

    ## --- "Summary & Intensity" sheet: current-year scope totals ---
    try:
        df = wb.get("Summary & Intensity", pd.DataFrame()).astype(str)  ## Load as strings for keyword search
        for _, row in df.iterrows():
            rs = " ".join(row.fillna("").astype(str))             ## Combine row into searchable string
            nums = pd.to_numeric(row, errors="coerce").dropna()   ## Extract all numeric values from this row
            if "Scope 1" in rs and "Direct" in rs and len(nums):
                data["s1_current"] = float(nums.iloc[0])          ## Scope 1 total: first number on the row
            if "Scope 2" in rs and "Market" in rs and "PRIMARY" in rs and len(nums):
                data["s2_mb_current"] = float(nums.iloc[0])       ## Scope 2 market-based: row marked PRIMARY
            if "Scope 2" in rs and "Location" in rs and "SECONDARY" in rs and len(nums):
                data["s2_lb_current"] = float(nums.iloc[0])       ## Scope 2 location-based: row marked SECONDARY
            if "Scope 3" in rs and "Value Chain" in rs and len(nums):
                data["s3_current"] = float(nums.iloc[0])          ## Scope 3 total: row mentioning "Value Chain"
    except Exception: pass

    ## --- "Reduction Trajectory" sheet: 5-year historical data by scope ---
    try:
        df = wb.get("Reduction Trajectory", pd.DataFrame())   ## Load raw (no header assumption)
        raw = df.astype(str)                                   ## String version for keyword detection
        yr_r = s1r = s2lbr = s2mbr = s3r = totr = None       ## Row index variables, all start as None

        for i, row in raw.iterrows():  ## Scan every row to find labeled data rows
            rs = " ".join(row.fillna("").astype(str))
            if "Scope / Year" in rs: yr_r = i          ## The row containing year column headers (2020, 2021, ...)
            if "Scope 1" in rs and "Direct" in rs: s1r = i        ## Row with annual Scope 1 totals
            if "location based" in rs.lower() and "Scope 2" in rs: s2lbr = i  ## Scope 2 location-based row
            if "market based" in rs.lower() and "Scope 2" in rs: s2mbr = i    ## Scope 2 market-based row
            if "Scope 3" in rs and "Value Chain" in rs: s3r = i   ## Scope 3 row
            if "TOTAL" in rs and "S1+S2+S3" in rs: totr = i       ## Combined S1+S2+S3 total row

        if yr_r is not None:  ## Only extract data if the year header row was found
            def extract(ridx):
                ## Build a {year: value} dict by pairing the data row with the year header row
                if ridx is None: return {}    ## Return empty dict if the row label was never found
                vals = pd.to_numeric(df.iloc[ridx], errors="coerce")  ## Parse data row as numbers
                yv   = pd.to_numeric(df.iloc[yr_r], errors="coerce")  ## Parse year header row as numbers
                return {int(yv.iloc[c]): float(vals.iloc[c])           ## {year: emission_value} pairs
                        for c in range(len(vals))
                        if not pd.isna(yv.iloc[c]) and not pd.isna(vals.iloc[c])  ## Skip columns missing either value
                        and 2015 <= yv.iloc[c] <= 2035}                ## Only include years in a valid range

            data["hist_s1"]    = extract(s1r)    ## Historical Scope 1 by year
            data["hist_s2lb"]  = extract(s2lbr)  ## Historical Scope 2 location-based by year
            data["hist_s2mb"]  = extract(s2mbr)  ## Historical Scope 2 market-based by year
            data["hist_s3"]    = extract(s3r)    ## Historical Scope 3 by year
            data["hist_total"] = extract(totr)   ## Historical combined total by year
            yv = pd.to_numeric(df.iloc[yr_r], errors="coerce").dropna()       ## Get all year values
            data["hist_years"] = sorted([int(y) for y in yv if 2015 <= y <= 2035])  ## Sorted valid year list
    except Exception: pass

    ## --- "Reduction Trajectory" sheet: future scenario pathways ---
    try:
        df = wb.get("Reduction Trajectory", pd.DataFrame())  ## Same sheet, scanning for trajectory rows
        raw = df.astype(str)
        bau_r = nz_r = sbti_r = s3t_r = ty_r = None  ## Row index variables for each scenario

        for i, row in raw.iterrows():
            rs = " ".join(row.fillna("").astype(str))
            if "Business as usual" in rs: bau_r = i             ## BAU: emissions grow at historical CAGR
            if "Net Zero Commitment" in rs and bau_r and not nz_r: nz_r = i  ## Net Zero: linear decline to target (must come after BAU row)
            if "Near-Term Science-Based" in rs: sbti_r = i      ## SBTi Scope 1+2 near-term target
            if "Scope 3 Reduction Target" in rs: s3t_r = i      ## SBTi Scope 3 value chain target
            if "Category / Year" in rs: ty_r = i                ## Year header row for the trajectory table

        if ty_r is not None:  ## Only proceed if the trajectory year header was found
            def ext_t(ridx):
                ## Same pattern as extract() above but uses the trajectory table's year header
                if ridx is None: return {}
                vals = pd.to_numeric(df.iloc[ridx], errors="coerce")
                yv   = pd.to_numeric(df.iloc[ty_r], errors="coerce")
                return {int(yv.iloc[c]): float(vals.iloc[c])
                        for c in range(len(vals))
                        if not pd.isna(yv.iloc[c]) and not pd.isna(vals.iloc[c])
                        and 2020 <= yv.iloc[c] <= 2035}  ## Trajectory covers 2020–2035

            data["traj_bau"]  = ext_t(bau_r)   ## Business-as-usual projection
            data["traj_nz"]   = ext_t(nz_r)    ## Net Zero pathway
            data["traj_sbti"] = ext_t(sbti_r)  ## SBTi Scope 1+2 target pathway
            data["traj_s3"]   = ext_t(s3t_r)   ## SBTi Scope 3 target pathway
    except Exception: pass

    return data  ## Return all extracted data as a dict

## ─── DEFAULT / DEMO DATA ─────────────────────────────────────────────────────
## Used when no file is uploaded and no manual values are entered
## Based on Microsoft's publicly reported GHG data 2020–2024
DEFAULT = {
    "company": "Your Company", "industry": "Technology",     ## Placeholder company info shown in header
    "revenue": 245100, "employees": 228000, "current_year": 2024,  ## Financial metrics: $M revenue, FTE headcount, reporting year
    "s1_current": 143510, "s2_mb_current": 259090,           ## Current year Scope 1 and Scope 2 market-based (tCO₂e)
    "s2_lb_current": 9955368, "s3_current": 15140000,        ## Current year Scope 2 location-based and Scope 3 (tCO₂e)
    "hist_years": [2020, 2021, 2022, 2023, 2024],            ## All years with historical data available
    ## Historical Scope 1 by year (tCO₂e) — direct combustion, company fleet, fugitive refrigerants
    "hist_s1":   {2020:118100,   2021:123704,  2022:139413,  2023:144960,  2024:143510},
    ## Historical Scope 2 market-based by year (tCO₂e) — purchased electricity adjusted for RECs/PPAs
    "hist_s2mb": {2020:456119,   2021:429405,  2022:288029,  2023:393134,  2024:259090},
    ## Historical Scope 2 location-based by year (tCO₂e) — purchased electricity using grid average emission factors
    "hist_s2lb": {2020:4328916,  2021:5010667, 2022:6381250, 2023:8077403, 2024:9955368},
    ## Historical Scope 3 by year (tCO₂e) — full upstream + downstream value chain
    "hist_s3":   {2020:11796000, 2021:13576000,2022:15916000,2023:16397000,2024:15140000},
    ## Historical total (S1 + S2 market-based + S3) by year (tCO₂e)
    "hist_total":{2020:12370000, 2021:14129000,2022:16343000,2023:16935000,2024:15543000},
    ## Future trajectory scenarios — used in the Reduction Trajectory tab
    "traj_bau":  {2024:12370000,2025:16456070,2026:17422778,2027:18446275,2028:19529898,2029:20677178,2030:21891855},  ## BAU: grows at historical CAGR
    "traj_nz":   {2024:12370000,2025:11906125,2026:11442250,2027:10978375,2028:10514500,2029:10050625,2030:9586750},   ## Net Zero: linear decline to 2030 target
    "traj_sbti": {2024:15543000,2025:12952500,2026:10362000,2027:7771500, 2028:5181000, 2029:2590500, 2030:0},        ## SBTi S1+2: full elimination by 2030
    "traj_s3":   {2024:11796000,2025:11206200,2026:10616400,2027:10026600,2028:9436800, 2029:8847000, 2030:8257200},  ## SBTi S3: 50% reduction by 2030 vs 2020 baseline
}

## ─── SIDEBAR ─────────────────────────────────────────────────────────────────
## Contains all data input controls: mode toggle, file uploader or manual entry, and target settings
with st.sidebar:
    ## App title rendered with Playfair Display serif for a consulting report feel
    st.markdown("""
    <div style="padding:20px 0 8px">
      <div style="font-family:Playfair Display,serif;font-size:1.2rem;font-weight:700;
                  color:white;line-height:1.2">GHG Emissions<br>Intelligence</div>
      <div style="font-size:0.7rem;letter-spacing:0.1em;text-transform:uppercase;
                  color:#99bbdd;margin-top:4px">GHG Protocol Framework</div>
    </div>""", unsafe_allow_html=True)
    st.divider()  ## Thin horizontal rule below the title

    ## Radio button lets user switch between two data input modes
    mode = st.radio("Data Source", ["📤 Upload Excel", "✏️ Manual Entry"])

    ## Start with a fresh copy of demo defaults on every Streamlit rerun
    ## Values are overwritten below by parsed or manually entered data
    d = dict(DEFAULT)

    if mode == "📤 Upload Excel":
        uploaded = st.file_uploader("Lab1_dashboard.xlsx", type=["xlsx"])  ## Accept only .xlsx files

        if uploaded:
            ## Build a unique ID from filename + file size to detect when a new file is uploaded
            ## Without this, parse_excel() would run on every Streamlit rerun (every click)
            file_id = f"{uploaded.name}_{uploaded.size}"

            if st.session_state.get("_file_id") != file_id:
                ## Re-parse only when the file changes — cached result persists in session_state
                with st.spinner("Parsing…"):
                    parsed = parse_excel(uploaded)              ## Run the full Excel parser
                    st.session_state["_parsed"] = parsed        ## Store result so it survives page reruns
                    st.session_state["_file_id"] = file_id      ## Remember which file was parsed

            ## Merge all non-falsy parsed values into d, overriding the defaults
            d.update({k: v for k, v in st.session_state["_parsed"].items() if v})
            st.success(f"✅ Loaded: {d.get(chr(39)+'company'+chr(39), 'file')}")  ## Confirm company name was loaded

        elif "_parsed" in st.session_state:
            ## Uploader is empty but a file was previously parsed — preserve that data
            ## This prevents the dashboard reverting to defaults when the user clicks a chart
            d.update({k: v for k, v in st.session_state["_parsed"].items() if v})
            st.caption(f"Using: {d.get(chr(39)+'company'+chr(39), 'uploaded file')}")

        else:
            ## No file uploaded and no session cache — show demo data notice
            st.caption("Demo data: Microsoft 2020–2024")

    else:
        ## ── MANUAL ENTRY MODE ──────────────────────────────────────────────────────
        ## All inputs write directly into d, which drives the entire dashboard

        st.markdown("**Company**")
        d["company"]       = st.text_input("Name", d["company"])       ## Company name shown in the page header
        d["industry"]      = st.text_input("Industry", d["industry"])  ## Industry shown below the company name
        d["current_year"]  = st.number_input("Reporting Year", value=int(d["current_year"]), step=1)   ## Controls which year is "current"
        d["revenue"]       = st.number_input("Revenue ($M)", value=float(d["revenue"]), step=1000.0)    ## Revenue in USD millions — used for intensity ratios
        d["employees"]     = st.number_input("Employees (FTE)", value=float(d["employees"]), step=1000.0)  ## Headcount — used for per-employee intensity

        st.markdown("**Current Year Emissions (tCO₂e)**")
        d["s1_current"]    = st.number_input("Scope 1", value=float(d["s1_current"]), step=1000.0)              ## Direct: stationary combustion, company vehicles, refrigerant leaks
        d["s2_mb_current"] = st.number_input("Scope 2 (Market-Based)", value=float(d["s2_mb_current"]), step=1000.0)    ## Purchased electricity, adjusted for RECs and PPAs — PRIMARY method per GHG Protocol
        d["s2_lb_current"] = st.number_input("Scope 2 (Location-Based)", value=float(d["s2_lb_current"]), step=1000.0)  ## Purchased electricity using grid average factors — SUPPLEMENTAL method
        d["s3_current"]    = st.number_input("Scope 3", value=float(d["s3_current"]), step=100000.0)            ## Full value chain: suppliers, employee commuting, product use, end-of-life

        st.markdown("**Historical Data — Scope 1 (tCO₂e)**")
        cy_val = int(d["current_year"])                           ## Current year as integer for arithmetic
        hist_input_years = list(range(cy_val - 4, cy_val))       ## Generate the 4 prior years: e.g. [2020,2021,2022,2023] if current is 2024
        d["hist_years"] = hist_input_years + [cy_val]            ## Full 5-year list including the current year

        for yr in hist_input_years:
            s1_default = float(d["hist_s1"].get(yr, 0.0))        ## Pre-fill from existing data or default to 0
            d["hist_s1"][yr] = st.number_input(                   ## Create one input per prior year
                f"Scope 1 — {yr}", value=s1_default, step=1000.0,
                key=f"s1_{yr}")                                   ## Unique key required — Streamlit tracks each widget by key

        st.markdown("**Historical Data — Scope 2 Market-Based (tCO₂e)**")
        for yr in hist_input_years:
            s2_default = float(d["hist_s2mb"].get(yr, 0.0))      ## Pre-fill or default to 0
            d["hist_s2mb"][yr] = st.number_input(
                f"Scope 2 MB — {yr}", value=s2_default, step=1000.0, key=f"s2mb_{yr}")

        st.markdown("**Historical Data — Scope 2 Location-Based (tCO₂e)**")
        for yr in hist_input_years:
            s2lb_default = float(d["hist_s2lb"].get(yr, 0.0))    ## Pre-fill or default to 0
            d["hist_s2lb"][yr] = st.number_input(
                f"Scope 2 LB — {yr}", value=s2lb_default, step=100000.0, key=f"s2lb_{yr}")

        st.markdown("**Historical Data — Scope 3 (tCO₂e)**")
        for yr in hist_input_years:
            s3_default = float(d["hist_s3"].get(yr, 0.0))        ## Pre-fill or default to 0
            d["hist_s3"][yr] = st.number_input(
                f"Scope 3 — {yr}", value=s3_default, step=100000.0, key=f"s3_{yr}")

        ## Auto-compute historical totals: S1 + S2 market-based + S3 for each prior year
        ## This mirrors how total_cy is computed below, keeping all totals consistent
        for yr in hist_input_years:
            d["hist_total"][yr] = d["hist_s1"].get(yr, 0) + d["hist_s2mb"].get(yr, 0) + d["hist_s3"].get(yr, 0)

        ## Mirror current year inputs into the historical dicts so charts include the current year
        d["hist_s1"][cy_val]    = d["s1_current"]
        d["hist_s2mb"][cy_val]  = d["s2_mb_current"]
        d["hist_s2lb"][cy_val]  = d["s2_lb_current"]
        d["hist_s3"][cy_val]    = d["s3_current"]
        d["hist_total"][cy_val] = d["s1_current"] + d["s2_mb_current"] + d["s3_current"]  ## Current year total

    st.divider()  ## Visual separator before the targets section
    st.markdown("**Targets**")
    nz_year   = st.number_input("Net Zero Year", value=2030, step=1)                ## The year the company aims to reach Net Zero
    nz_target = st.number_input("Net Zero Target (tCO₂e)", value=8659000, step=100000)  ## The absolute emission target in tCO₂e

## ─── DERIVED METRICS ─────────────────────────────────────────────────────────
## Compute all summary statistics used across KPI cards, headlines, and charts

total_cy = d["s1_current"] + d["s2_mb_current"] + d["s3_current"]  ## Full GHG inventory: S1 + S2 market-based + S3
cy = d["current_year"]  ## Current reporting year (integer)
py = cy - 1             ## Prior year used for year-on-year delta calculations

## Take the 5 most recent historical years that are at or before the current reporting year
hist_yrs5 = sorted([y for y in d["hist_years"] if y <= cy])[-5:]

## Year-on-year % change for each scope vs. the prior year — powers the delta arrows on KPI cards
d1_s1  = pct_chg(d["s1_current"],    d["hist_s1"].get(py))   ## Scope 1 YoY change
d1_s2  = pct_chg(d["s2_mb_current"], d["hist_s2mb"].get(py)) ## Scope 2 MB YoY change
d1_s3  = pct_chg(d["s3_current"],    d["hist_s3"].get(py))   ## Scope 3 YoY change
d1_tot = pct_chg(total_cy,           d["hist_total"].get(py)) ## Total YoY change

## Carbon intensity ratios — total emissions divided by a business activity metric
int_rev = total_cy / d["revenue"]   if d["revenue"]   else 0  ## tCO₂e per $M revenue
int_emp = total_cy / d["employees"] if d["employees"] else 0  ## tCO₂e per FTE employee

## Scope percentage shares — used in the donut chart and insight headline
s3_pct = d["s3_current"] / total_cy * 100 if total_cy else 0  ## Scope 3 as % of total
s1_pct = d["s1_current"] / total_cy * 100 if total_cy else 0  ## Scope 1 as % of total
s2_pct = d["s2_mb_current"] / total_cy * 100 if total_cy else 0  ## Scope 2 MB as % of total

## ─── PAGE HEADER ─────────────────────────────────────────────────────────────
## Full-width banner showing company name, year, revenue, and employee count
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

## ─── KPI ROW ─────────────────────────────────────────────────────────────────
## Five metric cards showing the most important numbers — the executive at-a-glance view
st.markdown(insight_header(
    f"Scope 3 dominates — {s3_pct:.0f}% of total footprint driven by value chain",  ## Dynamic headline using computed s3_pct
    f"Total inventory: {fmt_m(total_cy)} tCO₂e  ·  Reporting year {cy}  ·  Market-based Scope 2"
), unsafe_allow_html=True)

k1, k2, k3, k4, k5 = st.columns(5)  ## Five equal-width columns for the KPI cards
k1.markdown(kpi_block("Scope 1 · Direct",       fmt_m(d["s1_current"]),    "tCO₂e", d1_s1),  unsafe_allow_html=True)  ## Direct combustion, fleet, refrigerants
k2.markdown(kpi_block("Scope 2 · Market-Based", fmt_m(d["s2_mb_current"]), "tCO₂e", d1_s2),  unsafe_allow_html=True)  ## Purchased electricity adjusted for RECs/PPAs
k3.markdown(kpi_block("Scope 3 · Value Chain",  fmt_m(d["s3_current"]),    "tCO₂e", d1_s3),  unsafe_allow_html=True)  ## Full upstream + downstream value chain
k4.markdown(kpi_block("Total S1 + S2 + S3",     fmt_m(total_cy),           "tCO₂e", d1_tot), unsafe_allow_html=True)  ## Full GHG inventory combined
k5.markdown(kpi_block("Carbon Intensity",        f"{int_rev:.1f}", "tCO₂e / $M revenue",
                       note=f"{int_emp:.1f} tCO₂e / employee"), unsafe_allow_html=True)  ## Revenue intensity + employee intensity as a note

st.markdown("<hr>", unsafe_allow_html=True)  ## Horizontal rule separating KPIs from the tabs

## ─── TABS ────────────────────────────────────────────────────────────────────
## Four analytical views organized as tabs — each tells a different story about the data
tab1, tab2, tab3, tab4 = st.tabs([
    "Emission Trend",        ## 5-year absolute emissions history + year-on-year % change
    "Reduction Trajectory",  ## Historical actuals vs. BAU / Net Zero / SBTi future pathways
    "Carbon Intensity",      ## Intensity ratios over time + Scope 2 dual reporting comparison
    "Scope Breakdown"        ## Donut chart of current mix + stacked area showing 5-year composition
])

## ── TAB 1: EMISSION TREND ────────────────────────────────────────────────────
with tab1:
    ## Build 5-year value lists for each scope from the historical data dictionaries
    s1v = [d["hist_s1"].get(y)    for y in hist_yrs5]  ## Scope 1 for each of the 5 years
    s2v = [d["hist_s2mb"].get(y)  for y in hist_yrs5]  ## Scope 2 market-based for each year
    s3v = [d["hist_s3"].get(y)    for y in hist_yrs5]  ## Scope 3 for each year
    tv  = [d["hist_total"].get(y) for y in hist_yrs5]  ## Combined total for each year

    st.markdown(insight_header(
        "Absolute emissions declined in 2024 after peaking in 2023",
        "Scope 3 value chain emissions are the primary lever — representing over 97% of inventory"
    ), unsafe_allow_html=True)

    col_bar, col_yoy = st.columns([3, 2])  ## Left column (wider) for bar chart; right column for YoY chart

    with col_bar:
        fig = go.Figure()  ## New empty Plotly figure
        bw = 0.22          ## Bar width — three bars per year fit side by side at this width

        ## Scope 1 bars — shifted left by bw so they don't overlap with S2 and S3 bars
        fig.add_trace(go.Bar(name="Scope 1", x=[y-bw for y in hist_yrs5], y=s1v,
                             marker_color=C["s1"], width=bw,
                             text=[fmt_m(v) for v in s1v], textposition="outside",   ## Value labels above bars
                             textfont=dict(size=9, color=C["navy"])))

        ## Scope 2 bars — centered on the year position
        fig.add_trace(go.Bar(name="Scope 2 (Market)", x=hist_yrs5, y=s2v,
                             marker_color=C["s2"], width=bw,
                             text=[fmt_m(v) for v in s2v], textposition="outside",
                             textfont=dict(size=9, color=C["sky"])))

        ## Scope 3 bars — shifted right by bw
        fig.add_trace(go.Bar(name="Scope 3", x=[y+bw for y in hist_yrs5], y=s3v,
                             marker_color=C["s3"], width=bw,
                             text=[fmt_m(v) for v in s3v], textposition="outside",
                             textfont=dict(size=9, color=C["teal"])))

        ## Total as a dotted line with markers — overlaid on top of all bars for easy comparison
        fig.add_trace(go.Scatter(name="Total", x=hist_yrs5, y=tv,
                                 mode="lines+markers+text",
                                 text=[fmt_m(v) for v in tv], textposition="top center",  ## Labels above the dots
                                 textfont=dict(size=9, color=C["total"], family="IBM Plex Mono"),
                                 line=dict(color=C["total"], width=2, dash="dot"),          ## Dotted to distinguish from bars
                                 marker=dict(size=7, color=C["total"],
                                             line=dict(color="white", width=2))))           ## White ring makes markers pop

        apply_layout(fig, title="Absolute Emissions by Scope (tCO₂e)", height=420,
                     xaxis=dict(tickvals=hist_yrs5, tickformat="d", showgrid=False,         ## Show only the 5 year ticks
                                zeroline=False, showline=True, linecolor="#d0d5dd",
                                tickfont=dict(size=10, color="#667085")),
                     yaxis=dict(showgrid=True, gridcolor="#EBEBEB", zeroline=False,
                                tickfont=dict(size=10, color="#667085"), tickformat=".2s"))  ## SI format: 15M, 300K
        st.plotly_chart(fig, use_container_width=True)  ## Render at full column width

    with col_yoy:
        ## Calculate year-on-year % change for each scope — first entry is None (no prior year to compare)
        yoy_s1  = [None]+[pct_chg(s1v[i], s1v[i-1])  for i in range(1, len(s1v))]   ## S1 YoY % change per year
        yoy_s3  = [None]+[pct_chg(s3v[i], s3v[i-1])  for i in range(1, len(s3v))]   ## S3 YoY % change per year
        yoy_tot = [None]+[pct_chg(tv[i],  tv[i-1])    for i in range(1, len(tv))]    ## Total YoY % change per year

        fig2 = go.Figure()
        fig2.add_hline(y=0, line_color="#d0d5dd", line_width=1.5)  ## Zero reference line — above = growth, below = reduction

        ## Scope 1 YoY as dotted line with small markers (context, not the hero metric)
        fig2.add_trace(go.Scatter(name="Scope 1 YoY", x=hist_yrs5, y=yoy_s1,
                                  mode="lines+markers", line=dict(color=C["s1"], width=1.5, dash="dot"),
                                  marker=dict(size=5)))

        ## Scope 3 YoY as dotted line with small markers
        fig2.add_trace(go.Scatter(name="Scope 3 YoY", x=hist_yrs5, y=yoy_s3,
                                  mode="lines+markers", line=dict(color=C["s3"], width=1.5, dash="dot"),
                                  marker=dict(size=5)))

        ## Total YoY as solid line with value labels — the primary story line
        fig2.add_trace(go.Scatter(name="Total YoY", x=hist_yrs5, y=yoy_tot,
                                  mode="lines+markers+text",
                                  text=[f"{v:.1f}%" if v else "" for v in yoy_tot],  ## Show % label only when value exists
                                  textposition="top center",
                                  textfont=dict(size=9, color=C["total"]),
                                  line=dict(color=C["total"], width=2),
                                  marker=dict(size=7, color=C["total"],
                                              line=dict(color="white", width=2))))  ## White ring marker

        apply_layout(fig2, title="Year-on-Year Change (%)", height=420,
                     xaxis=dict(tickvals=hist_yrs5, tickformat="d", showgrid=False,
                                zeroline=False, showline=True, linecolor="#d0d5dd",
                                tickfont=dict(size=10, color="#667085")),
                     yaxis=dict(showgrid=True, gridcolor="#EBEBEB", zeroline=False,
                                tickfont=dict(size=10, color="#667085"), ticksuffix="%"))  ## Append % to y-axis labels
        st.plotly_chart(fig2, use_container_width=True)

    ## Collapsible data table beneath the charts — useful for detailed review or copy-paste
    with st.expander("View Data Table"):
        tbl = pd.DataFrame({"Year": hist_yrs5, "Scope 1": s1v, "Scope 2 MB": s2v, "Scope 3": s3v, "Total": tv})
        st.dataframe(tbl.style.format(
            {c: "{:,.0f}" for c in tbl.columns if c != "Year"}),  ## Comma format all numeric cols except Year
            use_container_width=True, hide_index=True)

## ── TAB 2: REDUCTION TRAJECTORY ──────────────────────────────────────────────
with tab2:
    bau_val = d.get("traj_bau", {}).get(nz_year, 0)   ## BAU projected value in the Net Zero target year
    gap = (bau_val - nz_target) / 1e6                  ## Gap between BAU and target, expressed in millions of tCO₂e

    st.markdown(insight_header(
        f"A {gap:.0f}M tCO₂e gap separates business-as-usual from the {nz_year} Net Zero target",
        "Accelerated decarbonisation across all scopes is required to meet SBTi-aligned pathways"
    ), unsafe_allow_html=True)

    fig3 = go.Figure()  ## New figure for the trajectory comparison chart

    ## Historical actuals plotted as a thick navy solid line — the known data
    fig3.add_trace(go.Scatter(
        name="Historical Actual", x=hist_yrs5,
        y=[d["hist_total"].get(y) for y in hist_yrs5],  ## Total emissions for each historical year
        mode="lines+markers", line=dict(color=C["navy"], width=3),
        marker=dict(size=8, color=C["navy"], line=dict(color="white", width=2)),  ## White-ringed markers
    ))

    def add_line(name, key, color, dash, width=1.5):
        ## Helper to add a scenario pathway line to fig3
        ## key is a trajectory dict key in d (e.g. "traj_bau", "traj_nz")
        td = d.get(key, {})     ## Get the {year: value} scenario dict
        if not td: return        ## Skip if data is missing (e.g. not yet parsed from Excel)
        yrs = sorted(td.keys()) ## Sort years chronologically
        fig3.add_trace(go.Scatter(
            name=name, x=yrs, y=[td[y] for y in yrs],
            mode="lines",  ## No markers on scenario lines — keeps the chart clean
            line=dict(color=color, width=width, dash=dash),
        ))

    add_line("Business As Usual",   "traj_bau",  C["bau"], "dash")     ## BAU: dashed red — emissions keep growing
    add_line("Net Zero Pathway",    "traj_nz",   C["nz"],  "solid", 2) ## Net Zero: solid teal — ideal trajectory
    add_line("SBTi Scope 1+2",      "traj_sbti", C["sky"], "dot")      ## SBTi S1+2 near-term target: dotted sky blue
    add_line("SBTi Scope 3 Target", "traj_s3",   "#5B5EA6","dot")      ## SBTi S3 value chain target: dotted indigo

    ## Callout annotation pointing to the Net Zero target on the chart
    fig3.add_annotation(
        x=nz_year, y=nz_target,                                          ## Anchor at the target year and emission level
        text=f"Net Zero Target<br><b>{fmt_m(nz_target)} tCO₂e</b>",     ## Two-line label: title + formatted value
        showarrow=True, arrowhead=2, arrowcolor=C["nz"], arrowwidth=1.5, ## Teal arrowhead
        font=dict(color=C["nz"], size=10, family="IBM Plex Sans"),
        bgcolor="white", bordercolor=C["nz"], borderwidth=1, borderpad=6,  ## White box, teal border
        ax=80, ay=-60  ## Arrow offset: label is 80px right and 60px above the data point
    )
    apply_layout(fig3, title="Emission Trajectory vs. Reduction Pathways (tCO₂e)", height=460,
                 xaxis=dict(tickformat="d", showgrid=False, zeroline=False,
                            showline=True, linecolor="#d0d5dd",
                            tickfont=dict(size=10, color="#667085")),
                 yaxis=dict(showgrid=True, gridcolor="#EBEBEB", zeroline=False,
                            tickfont=dict(size=10, color="#667085"), tickformat=".2s"))
    st.plotly_chart(fig3, use_container_width=True)

    ## Targets summary table below the trajectory chart
    st.markdown(section_rule("Reduction Targets"), unsafe_allow_html=True)
    st.dataframe(pd.DataFrame([
        {"Target": "Net Zero Commitment",     "Year": nz_year, "Baseline (tCO₂e)": "12,370,000",
         "Reduction Required": "30%",  "Target (tCO₂e)": f"{nz_target:,.0f}", "Framework": "GHG Protocol"},  ## Uses sidebar-configured nz_target
        {"Target": "SBTi Near-Term (S1+S2)", "Year": 2030,    "Baseline (tCO₂e)": "15,543,000",
         "Reduction Required": "100%", "Target (tCO₂e)": "0",             "Framework": "SBTi 1.5°C"},        ## Full elimination of S1+S2 by 2030
        {"Target": "SBTi Scope 3 Reduction", "Year": 2030,    "Baseline (tCO₂e)": "11,796,000",
         "Reduction Required": "50%",  "Target (tCO₂e)": "5,898,000",    "Framework": "SBTi value chain"},   ## 50% reduction from 2020 baseline
    ]), use_container_width=True, hide_index=True)

## ── TAB 3: CARBON INTENSITY ───────────────────────────────────────────────────
with tab3:
    ## Calculate prior year revenue intensity to determine the direction and magnitude of change
    int_rev_prev = d["hist_total"].get(py, 0) / d["revenue"] if d["revenue"] else 0  ## Prior year tCO₂e per $M
    int_chg = pct_chg(int_rev, int_rev_prev)                                          ## YoY % change in revenue intensity
    dir_word = "improved" if int_chg and int_chg < 0 else "increased"                 ## "improved" = fell (good); "increased" = rose (bad)

    st.markdown(insight_header(
        (f"Carbon intensity {dir_word} {abs(int_chg):.1f}% — "
         f"now {int_rev:.1f} tCO₂e per $M revenue") if int_chg else "Carbon intensity ratios",
        "Intensity metrics decouple emissions performance from business growth"
    ), unsafe_allow_html=True)

    ## Build 5-year intensity series by dividing historical totals by current financial metrics
    ## Note: we use current revenue/employees as a proxy — for rigorous reporting, use year-specific values
    int_rev_vals = [d["hist_total"].get(y, 0) / d["revenue"]   if d["revenue"]   else None for y in hist_yrs5]  ## tCO₂e per $M, each year
    int_emp_vals = [d["hist_total"].get(y, 0) / d["employees"] if d["employees"] else None for y in hist_yrs5]  ## tCO₂e per FTE, each year

    ci1, ci2 = st.columns(2)  ## Two equal columns for the two intensity charts

    def intensity_chart(yvals, title, color):
        ## Build a single-series bar chart with direct value labels above each bar
        ## No legend needed — one series, one color, title explains what it is
        f = go.Figure()
        f.add_trace(go.Bar(
            x=hist_yrs5, y=yvals,              ## Year on x-axis, intensity value on y-axis
            marker_color=color, opacity=0.85,  ## Slightly transparent for softer look
            width=0.5,                         ## Wider bars (only one per year, plenty of space)
            text=[f"{v:.1f}" if v else "" for v in yvals],  ## One-decimal value label above each bar
            textposition="outside",
            textfont=dict(size=10, color=color, family="IBM Plex Mono")))  ## Monospace for number alignment
        apply_layout(f, title=title, height=360, showlegend=False,  ## No legend for single-series chart
                     margin=dict(t=52, b=40, l=64, r=20),
                     xaxis=dict(tickvals=hist_yrs5, tickformat="d", showgrid=False,
                                zeroline=False, showline=True, linecolor="#d0d5dd",
                                tickfont=dict(size=10, color="#667085")),
                     yaxis=dict(showgrid=True, gridcolor="#EBEBEB", zeroline=False,
                                tickfont=dict(size=10, color="#667085")))
        return f

    with ci1:
        st.plotly_chart(intensity_chart(int_rev_vals,
            "Carbon Intensity — Revenue (tCO₂e / $M)", C["navy"]), use_container_width=True)  ## Navy bars for revenue intensity

    with ci2:
        st.plotly_chart(intensity_chart(int_emp_vals,
            "Carbon Intensity — Employee (tCO₂e / FTE)", C["teal"]), use_container_width=True)  ## Teal bars for employee intensity

    ## Scope 2 dual reporting section — required by GHG Protocol
    st.markdown(section_rule("Scope 2: Location-Based vs. Market-Based"), unsafe_allow_html=True)
    st.markdown("""<div style="font-size:0.82rem;color:#667085;margin-bottom:16px;max-width:680px">
    GHG Protocol requires dual reporting. The gap between location-based and market-based figures
    reflects the impact of renewable energy certificates (RECs) and power purchase agreements (PPAs).
    </div>""", unsafe_allow_html=True)

    s2lb_v = [d["hist_s2lb"].get(y) for y in hist_yrs5]  ## Location-based Scope 2 by year (grid averages)
    s2mb_v = [d["hist_s2mb"].get(y) for y in hist_yrs5]  ## Market-based Scope 2 by year (REC/PPA adjusted)

    fig_s2 = go.Figure()
    ## Location-based bars — offset left, semi-transparent slate (secondary method per GHG Protocol)
    fig_s2.add_trace(go.Bar(name="Location-Based (grid average)",   x=[y-0.18 for y in hist_yrs5],
                            y=s2lb_v, marker_color=C["slate"], opacity=0.5, width=0.3,
                            text=[fmt_m(v) for v in s2lb_v], textposition="outside",
                            textfont=dict(size=9, color=C["slate"])))
    ## Market-based bars — centered, full opacity sky blue (primary method per GHG Protocol)
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

## ── TAB 4: SCOPE BREAKDOWN ───────────────────────────────────────────────────
with tab4:
    st.markdown(insight_header(
        f"Value chain (Scope 3) is the critical focus — {s3_pct:.0f}% of total footprint",  ## Dynamic % computed above
        f"Scope 1: {s1_pct:.1f}%  ·  Scope 2 (market-based): {s2_pct:.1f}%  ·  Scope 3: {s3_pct:.1f}%"
    ), unsafe_allow_html=True)

    col_pie, col_area = st.columns(2)  ## Donut chart on the left, stacked area chart on the right

    with col_pie:
        ## Donut chart showing the current-year scope mix as percentage slices
        fig_pie = go.Figure(go.Pie(
            labels=["Scope 1 — Direct", "Scope 2 — Market", "Scope 3 — Value Chain"],  ## Three slice labels
            values=[d["s1_current"], d["s2_mb_current"], d["s3_current"]],              ## Current year tCO₂e per scope
            hole=0.62,  ## Large hole creates the donut; total is displayed as a center annotation
            marker=dict(
                colors=[C["s1"], C["s2"], C["s3"]],      ## Navy / sky / teal for the three scopes
                line=dict(color="white", width=3)),        ## White gaps between slices for visual separation
            textinfo="label+percent",                      ## Show both the label and the percentage on each slice
            textfont=dict(family="IBM Plex Sans", size=11, color="#2D3748"),
            direction="clockwise", sort=False,             ## Keep S1/S2/S3 in order; don't auto-sort by size
        ))
        ## Center annotation showing the total tCO₂e inside the donut hole
        fig_pie.add_annotation(
            text=(f"<b>{fmt_m(total_cy)}</b><br>"
                  f"<span style='font-size:11px;color:#667085'>tCO₂e</span>"),
            x=0.5, y=0.5, showarrow=False,  ## Position at the center of the chart
            font=dict(color="#003366", size=14, family="Playfair Display")
        )
        fig_pie.update_layout(
            title=f"Scope Mix — {cy}", height=420,
            paper_bgcolor="white", plot_bgcolor="white",
            font=dict(family="IBM Plex Sans", color="#2D3748"),
            legend=dict(bgcolor="white", bordercolor="#d0d5dd", borderwidth=1,
                        font=dict(size=10), orientation="h", yanchor="top", y=-0.05),  ## Legend below the chart
            margin=dict(t=52, b=40, l=20, r=20)
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    with col_area:
        ## Stacked area chart showing how each scope's absolute contribution has evolved over 5 years
        fig_area = go.Figure()

        ## Scope 3 goes first — it's the largest and forms the base of the stack (largest area at bottom)
        fig_area.add_trace(go.Scatter(
            name="Scope 3", x=hist_yrs5,
            y=[d["hist_s3"].get(y, 0) for y in hist_yrs5],  ## S3 historical values
            stackgroup="one",                                  ## All traces in "one" are stacked on top of each other
            fillcolor="rgba(0,107,107,0.75)",                 ## Semi-transparent teal fill
            line=dict(color=C["s3"], width=0)                 ## No visible border line
        ))
        ## Scope 2 is stacked on top of Scope 3
        fig_area.add_trace(go.Scatter(
            name="Scope 2 (Market)", x=hist_yrs5,
            y=[d["hist_s2mb"].get(y, 0) for y in hist_yrs5],
            stackgroup="one",                                  ## Same stackgroup — stacks above S3
            fillcolor="rgba(0,119,182,0.75)",
            line=dict(color=C["s2"], width=0)
        ))
        ## Scope 1 is stacked on top of Scope 2 (thinnest band at the top)
        fig_area.add_trace(go.Scatter(
            name="Scope 1", x=hist_yrs5,
            y=[d["hist_s1"].get(y, 0) for y in hist_yrs5],
            stackgroup="one",                                  ## Stacked above S2
            fillcolor="rgba(0,51,102,0.85)",                  ## Slightly more opaque navy at the top
            line=dict(color=C["s1"], width=0)
        ))
        apply_layout(fig_area, title="Stacked Composition Over Time (tCO₂e)", height=420,
                     xaxis=dict(tickvals=hist_yrs5, tickformat="d", showgrid=False,
                                zeroline=False, showline=True, linecolor="#d0d5dd",
                                tickfont=dict(size=10, color="#667085")),
                     yaxis=dict(showgrid=True, gridcolor="#EBEBEB", zeroline=False,
                                tickfont=dict(size=10, color="#667085"), tickformat=".2s"))
        st.plotly_chart(fig_area, use_container_width=True)

## ─── EXPORT SECTION ──────────────────────────────────────────────────────────
st.markdown("<hr>", unsafe_allow_html=True)              ## Separator above the export area
st.markdown(section_rule("Export"), unsafe_allow_html=True)  ## Section divider with navy rule

if st.button("Generate CSV Summary"):  ## Button click builds and shows the download button
    ## Build rows as a list of [label, value, unit] triplets
    rows = (
        [["CARBON FOOTPRINT SUMMARY REPORT", "", ""],                        ## Report title
         [f"Company: {d['company']}", f"Year: {cy}", f"Industry: {d['industry']}"],  ## Company context
         ["", "", ""],                                                        ## Blank spacer
         ["SCOPE TOTALS", "tCO₂e", ""],                                      ## Section header
         ["Scope 1 — Direct", d["s1_current"], ""],                          ## Current year S1
         ["Scope 2 — Market-Based", d["s2_mb_current"], ""],                 ## Current year S2 MB
         ["Scope 2 — Location-Based", d["s2_lb_current"], ""],               ## Current year S2 LB
         ["Scope 3 — Value Chain", d["s3_current"], ""],                     ## Current year S3
         ["Total (S1+S2MB+S3)", total_cy, ""],                               ## Current year combined total
         ["", "", ""],                                                        ## Blank spacer
         ["INTENSITY", "", ""],                                               ## Intensity section header
         ["Carbon Intensity — Revenue", round(int_rev, 2), "tCO₂e / $M"],   ## Revenue intensity, 2 decimals
         ["Carbon Intensity — Employee", round(int_emp, 2), "tCO₂e / FTE"], ## Employee intensity
         ["", "", ""],                                                        ## Blank spacer
         ["HISTORICAL TREND", "", ""]] +                                      ## Historical section header
        [[y, d["hist_total"].get(y, ""), "tCO₂e"] for y in hist_yrs5]       ## One row per historical year
    )
    csv = pd.DataFrame(rows).to_csv(index=False, header=False)  ## Convert to CSV string with no pandas index or headers
    st.download_button("⬇ Download CSV", csv,
                       file_name=f"GHG_Report_{d['company']}_{cy}.csv",  ## Dynamic filename: company + year
                       mime="text/csv")

## ─── FOOTER ──────────────────────────────────────────────────────────────────
## Attribution line with data source credits and reporting year, right-aligned
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
