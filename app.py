import streamlit as st
import pandas as pd
import json
import anthropic
from datetime import datetime, date
import io
from supabase import create_client, Client
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Orchestration-Digital",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Light Theme CSS ───────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

  html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    background-color: #ffffff !important;
    color: #1a1d23;
  }

  .stApp { background: #f8f9fc !important; }

  /* Sidebar */
  [data-testid="stSidebar"] {
    background: #ffffff !important;
    border-right: 1px solid #e8eaf0 !important;
  }
  [data-testid="stSidebar"] * { color: #1a1d23 !important; }

  /* Hero */
  .hero-header {
    background: linear-gradient(135deg, #1e3a5f 0%, #2d5a9b 50%, #1a8a6e 100%);
    border-radius: 16px;
    padding: 36px 40px;
    margin-bottom: 28px;
    position: relative;
    overflow: hidden;
  }
  .hero-header::after {
    content: '';
    position: absolute;
    top: -40px; right: -40px;
    width: 200px; height: 200px;
    background: rgba(255,255,255,0.05);
    border-radius: 50%;
  }
  .hero-badge {
    display: inline-block;
    background: rgba(255,255,255,0.15);
    color: #ffffff;
    border: 1px solid rgba(255,255,255,0.25);
    border-radius: 20px;
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    padding: 4px 14px;
    margin-bottom: 12px;
  }
  .hero-title {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 2rem;
    font-weight: 800;
    color: #ffffff;
    letter-spacing: -0.5px;
    margin: 0 0 8px 0;
  }
  .hero-sub { font-size: 0.9rem; color: rgba(255,255,255,0.75); margin: 0; }

  /* Cards */
  .metric-card {
    background: #ffffff;
    border: 1px solid #e8eaf0;
    border-radius: 12px;
    padding: 20px 24px;
    text-align: center;
    box-shadow: 0 1px 4px rgba(0,0,0,0.05);
  }
  .metric-value {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 1.7rem;
    font-weight: 700;
    color: #1e3a5f;
  }
  .metric-label {
    font-size: 0.75rem;
    color: #8a93a8;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin-top: 4px;
  }

  /* Section headers */
  .section-header {
    font-family: 'Plus Jakarta Sans', sans-serif;
    font-size: 1rem;
    font-weight: 700;
    color: #1a1d23;
    border-left: 3px solid #2d5a9b;
    padding-left: 12px;
    margin: 24px 0 14px 0;
  }

  /* Plan card */
  .plan-card {
    background: #ffffff;
    border: 1px solid #e8eaf0;
    border-radius: 14px;
    padding: 28px 32px;
    margin-top: 16px;
    line-height: 1.8;
    color: #2d3341;
    box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    font-size: 0.92rem;
  }

  /* Upload zone */
  [data-testid="stFileUploader"] {
    background: #f8f9fc !important;
    border: 2px dashed #c5cfe0 !important;
    border-radius: 12px !important;
  }

  /* Buttons */
  .stButton > button {
    background: linear-gradient(135deg, #1e3a5f, #2d5a9b) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
    padding: 10px 24px !important;
    transition: opacity 0.2s !important;
  }
  .stButton > button:hover { opacity: 0.88 !important; }

  /* Next button special */
  .next-btn .stButton > button {
    background: linear-gradient(135deg, #1a8a6e, #22b894) !important;
    width: 100%;
  }

  /* Inputs */
  .stTextInput > div > div > input,
  .stNumberInput > div > div > input,
  .stTextArea textarea,
  .stDateInput > div > div > input {
    background: #ffffff !important;
    border: 1px solid #d0d5e0 !important;
    border-radius: 8px !important;
    color: #1a1d23 !important;
    font-size: 0.9rem !important;
  }
  .stSelectbox > div > div {
    background: #ffffff !important;
    border: 1px solid #d0d5e0 !important;
    border-radius: 8px !important;
    color: #1a1d23 !important;
  }

  /* Tabs */
  .stTabs [data-baseweb="tab-list"] {
    background: #eef1f8;
    border-radius: 10px;
    padding: 4px;
    gap: 4px;
  }
  .stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: #6b7590 !important;
    border-radius: 7px !important;
    font-weight: 500 !important;
    font-size: 0.88rem !important;
  }
  .stTabs [aria-selected="true"] {
    background: #ffffff !important;
    color: #1e3a5f !important;
    font-weight: 600 !important;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08) !important;
  }

  /* Expanders */
  div[data-testid="stExpander"] {
    background: #ffffff !important;
    border: 1px solid #e8eaf0 !important;
    border-radius: 10px !important;
  }

  /* Success/info */
  .stAlert { border-radius: 10px !important; }

  /* Radio */
  .stRadio label { color: #1a1d23 !important; }

  /* Checkboxes */
  .stCheckbox label { color: #1a1d23 !important; }

  /* Labels */
  label, .stSelectbox label, .stTextInput label,
  .stNumberInput label, .stDateInput label,
  .stTextArea label { color: #4a5168 !important; font-size: 0.85rem !important; font-weight: 500 !important; }

  /* Dataframe */
  .stDataFrame { border-radius: 10px !important; }

  /* Hide Streamlit chrome */
  #MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Clients ───────────────────────────────────────────────────────────────────
@st.cache_resource
def get_supabase() -> Client:
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

def get_anthropic():
    return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])

# ── Helpers ───────────────────────────────────────────────────────────────────
def parse_uploaded_file(f) -> pd.DataFrame:
    name = f.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(f)
    elif name.endswith((".xlsx", ".xls")):
        return pd.read_excel(f)
    return None

def summarise_dataframe(df: pd.DataFrame) -> str:
    return "\n".join([
        f"Rows: {len(df)}, Columns: {len(df.columns)}",
        f"Columns: {', '.join(df.columns.tolist())}",
        "", "Sample data (first 10 rows):",
        df.head(10).to_string(index=False),
        "", "Numeric statistics:",
        df.describe().to_string()
    ])

def format_lkr(value: float) -> str:
    return f"LKR {value:,.0f}"

def generate_campaign_plan(client, data_summary, campaign_name, objective,
                            total_budget_lkr, start_date, end_date,
                            target_audience, channels, market, kpi_focus) -> str:
    duration_days = (end_date - start_date).days
    usd_rate = 320
    budget_usd = total_budget_lkr / usd_rate

    prompt = f"""You are a senior digital media planner with 15+ years of experience planning campaigns across Search, Social, Programmatic, and Video channels in Sri Lanka and Southeast Asia.

Based on the historical campaign performance data below, create a detailed digital media plan.

--- CAMPAIGN BRIEF ---
Campaign Name: {campaign_name}
Objective: {objective}
Total Budget: LKR {total_budget_lkr:,.0f} (approx. USD {budget_usd:,.0f} at 320 rate)
Flight Dates: {start_date.strftime('%d %b %Y')} – {end_date.strftime('%d %b %Y')} ({duration_days} days)
Target Audience: {target_audience}
Channels: {', '.join(channels)}
Market: {market}
Primary KPI: {kpi_focus}

--- HISTORICAL PERFORMANCE DATA ---
{data_summary}

--- YOUR TASK ---
Produce a complete media plan with these sections:

1. EXECUTIVE SUMMARY
   Strategic rationale (3–4 sentences).

2. CHANNEL MIX & BUDGET ALLOCATION
   For each channel provide:
   - Budget in LKR (amount and % of total)
   - Rationale from historical data
   - Expected reach/impressions
   - Flight dates

3. CHANNEL-BY-CHANNEL TACTICS
   For each channel:
   - Ad formats, targeting, bidding strategy, creative recommendations

4. KPI TARGETS & BENCHMARKS
   Realistic targets based on historical data:
   - Primary KPI: {kpi_focus}
   - Secondary KPIs (CTR, CPM, CPC, ROAS)
   - Weekly pacing milestones

5. OPTIMISATION ROADMAP
   Week-by-week actions.

6. RISK FLAGS
   Risks from historical data and mitigation strategies.

Use LKR for all budget figures. Be specific with numbers justified by the historical data."""

    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )
    return msg.content[0].text

# ── Excel Export in Template Format ──────────────────────────────────────────
def build_template_excel(brief: dict, plan_text: str) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Media Plan"

    # Colours
    navy   = "1E3A5F"
    teal   = "1A8A6E"
    white  = "FFFFFF"
    light  = "F0F4FA"
    grey   = "6B7590"
    border_col = "D0D5E0"

    def cell_style(ws, row, col, value="", bold=False, bg=None, fg="1A1D23",
                   align="left", size=10, border=True, num_fmt=None):
        c = ws.cell(row=row, column=col, value=value)
        c.font = Font(name="Inter", bold=bold, color=fg, size=size)
        if bg:
            c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        if border:
            thin = Side(style="thin", color=border_col)
            c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        if num_fmt:
            c.number_format = num_fmt
        return c

    def merge_style(ws, r1, c1, r2, c2, value="", bold=False, bg=None,
                    fg="1A1D23", align="left", size=10):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        c = ws.cell(row=r1, column=c1, value=value)
        c.font = Font(name="Inter", bold=bold, color=fg, size=size)
        if bg:
            c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        return c

    # Column widths
    col_widths = [22, 36, 18, 16, 14, 16, 16, 18, 18, 36, 14, 14, 10]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── HEADER BLOCK ─────────────────────────────────────────────────────────
    ws.row_dimensions[1].height = 40
    merge_style(ws, 1, 1, 1, 13,
                value="ORCHESTRATION-DIGITAL  |  MEDIA PLAN",
                bold=True, bg=navy, fg=white, align="center", size=14)

    headers_meta = [
        ("Date", brief.get("plan_date", datetime.today().strftime("%d %b %Y"))),
        ("Client", brief.get("client", "")),
        ("Brand", brief.get("brand", "")),
        ("Campaign Name", brief.get("campaign_name", "")),
        ("Campaign Start", brief.get("start_date", "").strftime("%d %b %Y") if brief.get("start_date") else ""),
        ("Campaign End",   brief.get("end_date",   "").strftime("%d %b %Y") if brief.get("end_date")   else ""),
        ("Plan Version", "V1.0"),
    ]
    for i, (label, val) in enumerate(headers_meta, 2):
        ws.row_dimensions[i].height = 20
        cell_style(ws, i, 1, label, bold=True, bg=light, fg=navy, align="right")
        merge_style(ws, i, 2, i, 5, value=val, fg="1A1D23")
        # blank remaining cols
        for c in range(6, 14):
            cell_style(ws, i, c, bg=white)

    # ── COLUMN HEADERS ────────────────────────────────────────────────────────
    hdr_row = 10
    ws.row_dimensions[hdr_row].height = 36
    headers = [
        "Channel / Placement", "Campaign Objective", "Audience",
        "Primary KPI Type", "Buying Rate", "Primary KPI",
        "Spendable (USD)", "Spendable (LKR)", "Billable (LKR)",
        "Audience Description", "Start Date", "End Date", "Days"
    ]
    for ci, h in enumerate(headers, 1):
        cell_style(ws, hdr_row, ci, h, bold=True, bg=navy, fg=white,
                   align="center", size=9)

    # ── CHANNEL DATA ──────────────────────────────────────────────────────────
    usd_rate   = 320
    commission = 0.10
    ssc_rate   = 0.025641
    vat_rate   = 0.18
    wht_rate   = 0.163

    total_lkr   = brief.get("total_budget", 0)
    channels    = brief.get("channels", [])
    start_dt    = brief.get("start_date", date.today())
    end_dt      = brief.get("end_date",   date.today())
    days        = (end_dt - start_dt).days
    objective   = brief.get("objective", "")
    audience    = brief.get("target_audience", "")
    kpi         = brief.get("kpi_focus", "")

    # Simple budget split across channels
    n = len(channels) if channels else 1
    per_channel_lkr = round(total_lkr / n, 0)

    channel_groups = {
        "Meta": ["Facebook", "Instagram"],
        "Google": ["YouTube", "GDN", "Google Search"],
        "TikTok": ["TikTok In-Feed", "TikTok TopView"],
        "Programmatic": ["Display", "Video"],
        "LinkedIn": ["LinkedIn Feed"],
    }

    # map selected channels to group keys
    selected_groups = {}
    for ch in channels:
        for grp, _ in channel_groups.items():
            if grp.lower() in ch.lower():
                selected_groups[grp] = per_channel_lkr
                break

    channel_totals = {}
    current_row = hdr_row + 1

    def write_group_header(row, name, bg_col):
        ws.row_dimensions[row].height = 22
        merge_style(ws, row, 1, row, 13, value=name,
                    bold=True, bg=bg_col, fg=white, align="left", size=10)

    group_colours = {
        "Meta": "1877F2",
        "TikTok": "010101",
        "Google": "1A8A6E",
        "Programmatic": "7B5EA7",
        "LinkedIn": "0A66C2",
    }

    for grp_name, grp_lkr in selected_groups.items():
        write_group_header(current_row, grp_name.upper(), group_colours.get(grp_name, navy))
        current_row += 1
        channel_totals[grp_name] = 0

        sub_channels = channel_groups.get(grp_name, [grp_name])
        sub_n = len(sub_channels)
        sub_lkr = round(grp_lkr / sub_n, 0)
        sub_usd = round(sub_lkr / usd_rate, 2)
        billable = round(sub_lkr * 1.05, 0)

        for sub in sub_channels:
            ws.row_dimensions[current_row].height = 20
            data = [
                sub, objective, audience, kpi, "CPM",
                "", sub_usd, sub_lkr, billable,
                audience,
                start_dt.strftime("%d/%m/%Y"),
                end_dt.strftime("%d/%m/%Y"),
                days
            ]
            for ci, val in enumerate(data, 1):
                fmt = None
                if ci in (7, 8, 9):
                    fmt = '#,##0.00' if ci == 7 else '#,##0'
                cell_style(ws, current_row, ci, val,
                           bg=light if current_row % 2 == 0 else white,
                           align="center" if ci > 5 else "left",
                           num_fmt=fmt)
            channel_totals[grp_name] += sub_lkr
            current_row += 1

    # blank row
    current_row += 1

    # ── SUMMARY TABLE ─────────────────────────────────────────────────────────
    total_working = sum(channel_totals.values())
    agency_comm   = round(total_working * commission, 2)
    sub1          = total_working + agency_comm
    ssc           = round(sub1 * ssc_rate, 2)
    sub2          = sub1 + ssc
    vat           = round(sub2 * vat_rate, 2)
    wht           = round(sub2 * wht_rate, 2)
    total_invest  = sub2 + vat

    summary_rows = []
    for grp_name, grp_total in channel_totals.items():
        summary_rows.append((f"{grp_name} Investment (incl. fees)", grp_total))
    summary_rows += [
        ("Total Working Investment (LKR)", total_working),
        ("Agency Commission (10%)", agency_comm),
        ("Sub Total", sub1),
        ("SSC Levy (2.5641%)", ssc),
        ("Sub Total", sub2),
        ("VAT (18%)", vat),
        ("Withholding Tax (16.3%)", wht),
        ("TOTAL INVESTMENT (LKR)", total_invest),
    ]

    for label, val in summary_rows:
        ws.row_dimensions[current_row].height = 22
        is_total = "TOTAL INVESTMENT" in label
        is_sub   = label.startswith("Sub Total") or label.startswith("Total Working")
        bg = navy if is_total else (light if is_sub else white)
        fg_col = white if is_total else "1A1D23"

        merge_style(ws, current_row, 1, current_row, 8,
                    value=label, bold=is_total or is_sub,
                    bg=bg, fg=fg_col, align="right")
        cell_style(ws, current_row, 9, val,
                   bold=is_total or is_sub,
                   bg=bg, fg=fg_col,
                   align="right", num_fmt='#,##0.00')
        for c in range(10, 14):
            cell_style(ws, current_row, c, bg=white)
        current_row += 1

    # ── PLAN TEXT SHEET ───────────────────────────────────────────────────────
    ws2 = wb.create_sheet("AI Plan")
    ws2.column_dimensions["A"].width = 120
    ws2.cell(row=1, column=1, value="AI-GENERATED CAMPAIGN PLAN").font = \
        Font(name="Inter", bold=True, size=13, color=navy)
    ws2.cell(row=2, column=1, value="")
    for i, line in enumerate(plan_text.split("\n"), 3):
        c = ws2.cell(row=i, column=1, value=line)
        c.font = Font(name="Inter", size=10)
        c.alignment = Alignment(wrap_text=True)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def save_plan_to_supabase(supabase, plan_data):
    try:
        supabase.table("campaign_plans").insert(plan_data).execute()
        return True
    except Exception as e:
        st.warning(f"Could not save: {e}")
        return False

def load_saved_plans(supabase):
    try:
        return supabase.table("campaign_plans").select("*").order("created_at", desc=True).limit(20).execute().data
    except:
        return []

# ── Session state defaults ────────────────────────────────────────────────────
if "active_tab" not in st.session_state:
    st.session_state["active_tab"] = 0

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='padding: 8px 0 24px 0;'>
        <div style='font-family: Plus Jakarta Sans, sans-serif; font-size: 1.25rem;
                    font-weight: 800; color: #1e3a5f;'>🎯 Orchestration</div>
        <div style='font-size: 0.75rem; color: #8a93a8; margin-top: 2px;'>Digital Planning Platform</div>
    </div>
    """, unsafe_allow_html=True)

    nav = st.radio("Navigation", ["📊 New Campaign Plan", "📁 Saved Plans", "📈 Data Explorer"],
                   label_visibility="collapsed")

    st.markdown("---")
    st.markdown("<div style='font-size:0.72rem; color:#8a93a8;'>POC Version 1.0 · Phase 1</div>",
                unsafe_allow_html=True)

# ── Hero ──────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero-header">
  <div class="hero-badge">Digital Planning Platform</div>
  <div class="hero-title">Orchestration-Digital</div>
  <div class="hero-sub">Upload past campaign data · Generate AI-powered media plans · Export in agency format</div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: NEW CAMPAIGN PLAN
# ══════════════════════════════════════════════════════════════════════════════
if nav == "📊 New Campaign Plan":

    tab_index = st.session_state.get("active_tab", 0)
    tab1, tab2, tab3 = st.tabs(["  1 · Upload Data  ", "  2 · Campaign Brief  ", "  3 · Your Plan  "])

    # ── Tab 1: Upload ─────────────────────────────────────────────────────────
    with tab1:
        st.markdown('<div class="section-header">Upload Historical Campaign Data</div>',
                    unsafe_allow_html=True)
        st.caption("Upload one or more CSV/Excel exports from Meta, Google, TikTok, or any ad platform.")

        uploaded_files = st.file_uploader(
            "Drop your campaign files here",
            type=["csv", "xlsx", "xls"],
            accept_multiple_files=True,
            help="You can select multiple files at once — hold Ctrl/Cmd to multi-select"
        )

        if uploaded_files:
            all_dfs = []
            file_names = []
            for f in uploaded_files:
                df = parse_uploaded_file(f)
                if df is not None:
                    all_dfs.append(df)
                    file_names.append(f.name)

            if all_dfs:
                combined_df = pd.concat(all_dfs, ignore_index=True)
                st.session_state["uploaded_df"] = combined_df
                st.session_state["file_names"] = file_names

                st.success(f"✅ {len(uploaded_files)} file(s) loaded — {len(combined_df):,} total rows")

                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown(f'<div class="metric-card"><div class="metric-value">{len(all_dfs)}</div><div class="metric-label">Files Uploaded</div></div>', unsafe_allow_html=True)
                with col2:
                    st.markdown(f'<div class="metric-card"><div class="metric-value">{len(combined_df):,}</div><div class="metric-label">Total Rows</div></div>', unsafe_allow_html=True)
                with col3:
                    st.markdown(f'<div class="metric-card"><div class="metric-value">{len(combined_df.columns)}</div><div class="metric-label">Columns</div></div>', unsafe_allow_html=True)

                with st.expander("Preview combined data"):
                    st.dataframe(combined_df.head(20), use_container_width=True)

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown('<div class="next-btn">', unsafe_allow_html=True)
                if st.button("Continue to Campaign Brief →", use_container_width=True):
                    st.session_state["active_tab"] = 1
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

        else:
            st.markdown("""
            <div style='text-align:center; padding:48px; color:#8a93a8; background:#f8f9fc;
                        border-radius:12px; border:2px dashed #d0d5e0;'>
                <div style='font-size:2.5rem; margin-bottom:12px;'>📂</div>
                <div style='font-weight:600; color:#4a5168; margin-bottom:6px;'>No files uploaded yet</div>
                <div style='font-size:0.85rem;'>Export data from Meta Ads Manager, Google Ads, or TikTok Ads and upload above.<br>
                You can upload multiple files at once.</div>
            </div>
            """, unsafe_allow_html=True)

    # ── Tab 2: Brief ──────────────────────────────────────────────────────────
    with tab2:
        if "uploaded_df" not in st.session_state:
            st.warning("⬅️ Please upload your campaign data first in the **Upload Data** tab.")
        else:
            st.markdown('<div class="section-header">Campaign Details</div>', unsafe_allow_html=True)

            col_a, col_b = st.columns(2)
            with col_a:
                client_name   = st.text_input("Client Name", placeholder="e.g. Unilever Lanka")
                brand_name    = st.text_input("Brand Name", placeholder="e.g. Sunlight")
                campaign_name = st.text_input("Campaign Name", placeholder="e.g. Q3 2025 Brand Awareness")
                objective     = st.selectbox("Campaign Objective", [
                    "Brand Awareness", "Reach & Frequency", "Website Traffic",
                    "Lead Generation", "App Installs", "E-commerce / Conversions",
                    "Video Views", "Engagement"
                ])

            with col_b:
                target_audience = st.text_area("Target Audience",
                    placeholder="e.g. 25–45, urban professionals, interested in finance & tech",
                    height=100)
                kpi_focus = st.selectbox("Primary KPI", [
                    "ROAS (Return on Ad Spend)", "CPA (Cost per Acquisition)",
                    "CPL (Cost per Lead)", "CTR (Click-through Rate)",
                    "CPM (Cost per 1000 Impressions)", "CPC (Cost per Click)",
                    "Reach", "Video Completion Rate"
                ])
                market = st.text_input("Market / Region", value="Sri Lanka")

            st.markdown('<div class="section-header">Budget & Dates</div>', unsafe_allow_html=True)
            col_c, col_d1, col_d2 = st.columns([2, 1, 1])
            with col_c:
                total_budget_lkr = st.number_input(
                    "Total Budget (LKR)",
                    min_value=0,
                    value=1000000,
                    step=50000,
                    format="%d",
                    help="Enter budget in Sri Lankan Rupees"
                )
                st.caption(f"💰 {format_lkr(total_budget_lkr)}  ≈  USD {total_budget_lkr/320:,.0f}")
            with col_d1:
                start_date = st.date_input("Start Date", value=date.today())
            with col_d2:
                end_date = st.date_input("End Date")

            st.markdown('<div class="section-header">Channel Selection</div>', unsafe_allow_html=True)
            cols = st.columns(3)
            ch_meta         = cols[0].checkbox("Meta (Facebook/Instagram)", value=True)
            ch_google       = cols[1].checkbox("Google Ads", value=True)
            ch_tiktok       = cols[2].checkbox("TikTok Ads", value=True)
            ch_programmatic = cols[0].checkbox("Programmatic Display")
            ch_youtube      = cols[1].checkbox("YouTube")
            ch_linkedin     = cols[2].checkbox("LinkedIn")

            channels_selected = []
            if ch_meta:         channels_selected.append("Meta (Facebook & Instagram)")
            if ch_google:       channels_selected.append("Google Search & Display")
            if ch_tiktok:       channels_selected.append("TikTok Ads")
            if ch_programmatic: channels_selected.append("Programmatic Display")
            if ch_youtube:      channels_selected.append("YouTube")
            if ch_linkedin:     channels_selected.append("LinkedIn")

            st.markdown("---")
            if st.button("🚀 Generate Campaign Plan", use_container_width=True):
                if not campaign_name:
                    st.error("Please enter a campaign name.")
                elif end_date <= start_date:
                    st.error("End date must be after start date.")
                elif not channels_selected:
                    st.error("Please select at least one channel.")
                else:
                    st.session_state["brief"] = {
                        "client": client_name,
                        "brand": brand_name,
                        "campaign_name": campaign_name,
                        "objective": objective,
                        "total_budget": total_budget_lkr,
                        "start_date": start_date,
                        "end_date": end_date,
                        "target_audience": target_audience,
                        "channels": channels_selected,
                        "market": market,
                        "kpi_focus": kpi_focus,
                        "plan_date": datetime.today().strftime("%d %b %Y"),
                    }
                    with st.spinner("🤖 Analysing your data and building the media plan…"):
                        try:
                            ai_client = get_anthropic()
                            summary   = summarise_dataframe(st.session_state["uploaded_df"])
                            plan_text = generate_campaign_plan(
                                ai_client, summary, campaign_name, objective,
                                total_budget_lkr, start_date, end_date,
                                target_audience, channels_selected, market, kpi_focus
                            )
                            st.session_state["generated_plan"] = plan_text
                            st.session_state["active_tab"] = 2
                            st.success("✅ Plan generated!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error: {e}")

    # ── Tab 3: Plan ───────────────────────────────────────────────────────────
    with tab3:
        if "generated_plan" not in st.session_state:
            st.info("👆 Complete the brief in **2 · Campaign Brief** and hit Generate.")
        else:
            brief    = st.session_state.get("brief", {})
            duration = (brief.get("end_date", date.today()) - brief.get("start_date", date.today())).days

            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown(f'<div class="metric-card"><div class="metric-value">{format_lkr(brief.get("total_budget",0))}</div><div class="metric-label">Total Budget</div></div>', unsafe_allow_html=True)
            with col2:
                st.markdown(f'<div class="metric-card"><div class="metric-value">{duration} days</div><div class="metric-label">Campaign Duration</div></div>', unsafe_allow_html=True)
            with col3:
                st.markdown(f'<div class="metric-card"><div class="metric-value">{len(brief.get("channels",[]))}</div><div class="metric-label">Channels</div></div>', unsafe_allow_html=True)

            plan_html = st.session_state["generated_plan"].replace("\n", "<br>")
            st.markdown(f'<div class="plan-card">{plan_html}</div>', unsafe_allow_html=True)

            st.markdown("---")
            col_e1, col_e2, col_e3 = st.columns(3)

            with col_e1:
                st.download_button(
                    "📥 Download as TXT",
                    data=st.session_state["generated_plan"],
                    file_name=f"{brief.get('campaign_name','plan').replace(' ','_')}_plan.txt",
                    mime="text/plain",
                    use_container_width=True
                )

            with col_e2:
                excel_bytes = build_template_excel(brief, st.session_state["generated_plan"])
                st.download_button(
                    "📊 Download Media Plan (Excel)",
                    data=excel_bytes,
                    file_name=f"{brief.get('campaign_name','plan').replace(' ','_')}_MediaPlan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            with col_e3:
                if st.button("💾 Save to Library", use_container_width=True):
                    supabase = get_supabase()
                    save_plan_to_supabase(supabase, {
                        "campaign_name": brief.get("campaign_name"),
                        "objective":     brief.get("objective"),
                        "total_budget":  brief.get("total_budget"),
                        "start_date":    str(brief.get("start_date")),
                        "end_date":      str(brief.get("end_date")),
                        "channels":      json.dumps(brief.get("channels", [])),
                        "market":        brief.get("market"),
                        "kpi_focus":     brief.get("kpi_focus"),
                        "plan_text":     st.session_state["generated_plan"],
                        "created_at":    datetime.utcnow().isoformat()
                    })
                    st.success("Saved to your plan library!")

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: SAVED PLANS
# ══════════════════════════════════════════════════════════════════════════════
elif nav == "📁 Saved Plans":
    st.markdown('<div class="section-header">Saved Campaign Plans</div>', unsafe_allow_html=True)
    try:
        supabase = get_supabase()
        plans = load_saved_plans(supabase)
        if not plans:
            st.info("No saved plans yet. Generate your first plan and save it to the library.")
        else:
            for plan in plans:
                with st.expander(f"📋 {plan.get('campaign_name','Unnamed')} — {plan.get('market','')} — {format_lkr(plan.get('total_budget',0))}"):
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Objective", plan.get("objective","—"))
                    c2.metric("Budget", format_lkr(plan.get("total_budget",0)))
                    c3.metric("KPI", plan.get("kpi_focus","—"))
                    st.markdown(f"**Channels:** {plan.get('channels','')}")
                    st.markdown(f"**Dates:** {plan.get('start_date')} → {plan.get('end_date')}")
                    st.markdown("---")
                    st.markdown(plan.get("plan_text",""))
    except Exception as e:
        st.error(f"Could not load plans: {e}")

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: DATA EXPLORER
# ══════════════════════════════════════════════════════════════════════════════
elif nav == "📈 Data Explorer":
    st.markdown('<div class="section-header">Data Explorer</div>', unsafe_allow_html=True)
    if "uploaded_df" not in st.session_state:
        st.info("Upload campaign data in the **New Campaign Plan** section first.")
    else:
        df = st.session_state["uploaded_df"]
        names = st.session_state.get("file_names", ["uploaded file"])
        st.caption(f"Exploring: **{', '.join(names)}** — {len(df):,} rows")
        st.dataframe(df, use_container_width=True)

        numeric_cols = df.select_dtypes(include='number').columns.tolist()
        if len(numeric_cols) >= 2:
            st.markdown('<div class="section-header">Quick Chart</div>', unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            x_col = c1.selectbox("X axis", df.columns.tolist())
            y_col = c2.selectbox("Y axis", numeric_cols)
            chart_type = st.radio("Chart type", ["Line", "Bar", "Area"], horizontal=True)
            chart_df = df[[x_col, y_col]].dropna().set_index(x_col)
            if chart_type == "Line":   st.line_chart(chart_df)
            elif chart_type == "Bar":  st.bar_chart(chart_df)
            else:                      st.area_chart(chart_df)
