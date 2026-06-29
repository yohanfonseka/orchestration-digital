import streamlit as st
import pandas as pd
import json
import anthropic
from datetime import datetime, date
import io
from supabase import create_client, Client
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Orchestration-Digital",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
  html, body, [class*="css"] { font-family: 'Inter', sans-serif; background-color: #ffffff !important; color: #1a1d23; }
  .stApp { background: #f8f9fc !important; }
  [data-testid="stSidebar"] { background: #ffffff !important; border-right: 1px solid #e8eaf0 !important; }
  [data-testid="stSidebar"] * { color: #1a1d23 !important; }

  .hero-header {
    background: linear-gradient(135deg, #1e3a5f 0%, #2d5a9b 50%, #1a8a6e 100%);
    border-radius: 16px; padding: 36px 40px; margin-bottom: 28px; position: relative; overflow: hidden;
  }
  .hero-badge {
    display: inline-block; background: rgba(255,255,255,0.15); color: #fff;
    border: 1px solid rgba(255,255,255,0.25); border-radius: 20px;
    font-size: 0.7rem; font-weight: 600; letter-spacing: 0.1em;
    text-transform: uppercase; padding: 4px 14px; margin-bottom: 12px;
  }
  .hero-title { font-family: 'Plus Jakarta Sans', sans-serif; font-size: 2rem; font-weight: 800; color: #fff; letter-spacing: -0.5px; margin: 0 0 8px 0; }
  .hero-sub { font-size: 0.9rem; color: rgba(255,255,255,0.75); margin: 0; }

  .metric-card { background: #fff; border: 1px solid #e8eaf0; border-radius: 12px; padding: 20px 24px; text-align: center; box-shadow: 0 1px 4px rgba(0,0,0,0.05); }
  .metric-value { font-family: 'Plus Jakarta Sans', sans-serif; font-size: 1.6rem; font-weight: 700; color: #1e3a5f; }
  .metric-label { font-size: 0.75rem; color: #8a93a8; text-transform: uppercase; letter-spacing: 0.06em; margin-top: 4px; }

  .section-header { font-family: 'Plus Jakarta Sans', sans-serif; font-size: 1rem; font-weight: 700; color: #1a1d23; border-left: 3px solid #2d5a9b; padding-left: 12px; margin: 24px 0 14px 0; }

  .step-card { background: #fff; border: 1px solid #e8eaf0; border-radius: 14px; padding: 24px 28px; margin-bottom: 16px; box-shadow: 0 1px 4px rgba(0,0,0,0.05); }
  .step-badge { display: inline-block; background: #1e3a5f; color: #fff; border-radius: 50%; width: 28px; height: 28px; line-height: 28px; text-align: center; font-weight: 700; font-size: 0.85rem; margin-right: 10px; }
  .step-title { font-family: 'Plus Jakarta Sans', sans-serif; font-size: 1.05rem; font-weight: 700; color: #1e3a5f; display: inline; }
  .step-complete { background: #f0faf6; border-color: #1a8a6e; }
  .step-active { border-color: #2d5a9b; border-width: 2px; }

  .brand-card { background: #fff; border: 1px solid #e8eaf0; border-radius: 10px; padding: 16px 20px; margin-bottom: 8px; cursor: pointer; transition: all 0.2s; }
  .brand-card:hover { border-color: #2d5a9b; box-shadow: 0 2px 8px rgba(45,90,155,0.1); }
  .brand-card-selected { border-color: #2d5a9b; background: #f0f4fa; border-width: 2px; }

  .plan-card { background: #fff; border: 1px solid #e8eaf0; border-radius: 14px; padding: 28px 32px; margin-top: 16px; line-height: 1.8; color: #2d3341; box-shadow: 0 1px 4px rgba(0,0,0,0.05); font-size: 0.92rem; }

  .stButton > button { background: linear-gradient(135deg, #1e3a5f, #2d5a9b) !important; color: white !important; border: none !important; border-radius: 8px !important; font-weight: 600 !important; font-size: 0.9rem !important; padding: 10px 24px !important; transition: opacity 0.2s !important; }
  .stButton > button:hover { opacity: 0.88 !important; }
  .green-btn .stButton > button { background: linear-gradient(135deg, #1a8a6e, #22b894) !important; }

  .stTextInput > div > div > input, .stNumberInput > div > div > input,
  .stTextArea textarea, .stDateInput > div > div > input {
    background: #fff !important; border: 1px solid #d0d5e0 !important;
    border-radius: 8px !important; color: #1a1d23 !important; font-size: 0.9rem !important;
  }
  .stSelectbox > div > div { background: #fff !important; border: 1px solid #d0d5e0 !important; border-radius: 8px !important; color: #1a1d23 !important; }

  .stTabs [data-baseweb="tab-list"] { background: #eef1f8; border-radius: 10px; padding: 4px; gap: 4px; }
  .stTabs [data-baseweb="tab"] { background: transparent !important; color: #6b7590 !important; border-radius: 7px !important; font-weight: 500 !important; font-size: 0.88rem !important; }
  .stTabs [aria-selected="true"] { background: #fff !important; color: #1e3a5f !important; font-weight: 600 !important; box-shadow: 0 1px 4px rgba(0,0,0,0.08) !important; }

  div[data-testid="stExpander"] { background: #fff !important; border: 1px solid #e8eaf0 !important; border-radius: 10px !important; }
  .stAlert { border-radius: 10px !important; }
  label, .stSelectbox label, .stTextInput label, .stNumberInput label,
  .stDateInput label, .stTextArea label { color: #4a5168 !important; font-size: 0.85rem !important; font-weight: 500 !important; }
  #MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Clients ───────────────────────────────────────────────────────────────────
@st.cache_resource
def get_supabase() -> Client:
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

def get_anthropic_client():
    return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])

# ── Supabase helpers ──────────────────────────────────────────────────────────
def get_clients_list(sb):
    try:
        r = sb.table("clients").select("*").order("client_name").execute()
        return r.data or []
    except:
        return []

def get_brands_for_client(sb, client_id):
    try:
        r = sb.table("brands").select("*").eq("client_id", client_id).order("brand_name").execute()
        return r.data or []
    except:
        return []

def create_client_record(sb, client_name):
    try:
        r = sb.table("clients").insert({"client_name": client_name, "created_at": datetime.utcnow().isoformat()}).execute()
        return r.data[0] if r.data else None
    except Exception as e:
        st.error(f"Error creating client: {e}")
        return None

def create_brand_record(sb, client_id, brand_name):
    try:
        r = sb.table("brands").insert({"client_id": client_id, "brand_name": brand_name, "created_at": datetime.utcnow().isoformat()}).execute()
        return r.data[0] if r.data else None
    except Exception as e:
        st.error(f"Error creating brand: {e}")
        return None

def get_brand_data_files(sb, brand_id):
    try:
        r = sb.table("brand_data").select("*").eq("brand_id", brand_id).order("uploaded_at", desc=True).execute()
        return r.data or []
    except:
        return []

def save_brand_data(sb, brand_id, file_name, data_json, row_count):
    try:
        sb.table("brand_data").insert({
            "brand_id": brand_id,
            "file_name": file_name,
            "data_json": data_json,
            "row_count": row_count,
            "uploaded_at": datetime.utcnow().isoformat()
        }).execute()
        return True
    except Exception as e:
        st.error(f"Error saving data: {e}")
        return False

def delete_brand_data_file(sb, file_id):
    try:
        sb.table("brand_data").delete().eq("id", file_id).execute()
        return True
    except:
        return False

def save_plan(sb, plan_data):
    try:
        sb.table("campaign_plans").insert(plan_data).execute()
        return True
    except Exception as e:
        st.warning(f"Could not save: {e}")
        return False

def load_saved_plans(sb, brand_id=None):
    try:
        q = sb.table("campaign_plans").select("*").order("created_at", desc=True)
        if brand_id:
            q = q.eq("brand_id", brand_id)
        return q.limit(30).execute().data or []
    except:
        return []

# ── Data helpers ──────────────────────────────────────────────────────────────
def parse_uploaded_file(f):
    name = f.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(f)
    elif name.endswith((".xlsx", ".xls")):
        return pd.read_excel(f)
    return None

def summarise_dataframe(df):
    return "\n".join([
        f"Rows: {len(df)}, Columns: {len(df.columns)}",
        f"Columns: {', '.join(df.columns.tolist())}",
        "", "Sample data (first 10 rows):",
        df.head(10).to_string(index=False),
        "", "Numeric statistics:",
        df.describe().to_string()
    ])

def format_lkr(value):
    return f"LKR {value:,.0f}"

# ── AI Plan Generator ─────────────────────────────────────────────────────────
def generate_campaign_plan(client, data_summary, brief):
    duration_days = (brief["end_date"] - brief["start_date"]).days
    budget_usd = brief["total_budget"] / 320

    prompt = f"""You are a senior digital media planner with 15+ years of experience planning campaigns in Sri Lanka and Southeast Asia.

Based on the historical campaign performance data below, create a detailed digital media plan.

--- CAMPAIGN BRIEF ---
Client: {brief["client_name"]}
Brand: {brief["brand_name"]}
Campaign Name: {brief["campaign_name"]}
Objective: {brief["objective"]}
Total Budget: LKR {brief["total_budget"]:,.0f} (approx. USD {budget_usd:,.0f})
Flight Dates: {brief["start_date"].strftime("%d %b %Y")} – {brief["end_date"].strftime("%d %b %Y")} ({duration_days} days)
Target Audience: {brief["target_audience"]}
Channels: {", ".join(brief["channels"])}
Market: {brief["market"]}
Primary KPI: {brief["kpi_focus"]}

--- HISTORICAL PERFORMANCE DATA ---
{data_summary}

--- YOUR TASK ---
Produce a complete media plan with these sections:

1. EXECUTIVE SUMMARY
   Strategic rationale (3–4 sentences).

2. CHANNEL MIX & BUDGET ALLOCATION
   For each channel:
   - Budget in LKR (amount and % of total)
   - Rationale from historical data
   - Expected reach/impressions
   - Flight dates

3. CHANNEL-BY-CHANNEL TACTICS
   Ad formats, targeting, bidding strategy, creative recommendations per channel.

4. KPI TARGETS & BENCHMARKS
   Realistic targets from historical data:
   - Primary KPI: {brief["kpi_focus"]}
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

# ── Excel Export ──────────────────────────────────────────────────────────────
def build_template_excel(brief, plan_text):
    wb = Workbook()
    ws = wb.active
    ws.title = "Media Plan"

    navy = "1E3A5F"; teal = "1A8A6E"; white = "FFFFFF"; light = "F0F4FA"; border_col = "D0D5E0"

    def cs(row, col, value="", bold=False, bg=None, fg="1A1D23", align="left", size=10, num_fmt=None):
        c = ws.cell(row=row, column=col, value=value)
        c.font = Font(name="Inter", bold=bold, color=fg, size=size)
        if bg: c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        thin = Side(style="thin", color=border_col)
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        if num_fmt: c.number_format = num_fmt
        return c

    def ms(r1, c1, r2, c2, value="", bold=False, bg=None, fg="1A1D23", align="left", size=10):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        c = ws.cell(row=r1, column=c1, value=value)
        c.font = Font(name="Inter", bold=bold, color=fg, size=size)
        if bg: c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        return c

    for i, w in enumerate([22,36,18,16,14,16,16,18,18,36,14,14,10], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 40
    ms(1,1,1,13, "ORCHESTRATION-DIGITAL  |  MEDIA PLAN", bold=True, bg=navy, fg=white, align="center", size=14)

    meta = [
        ("Date", brief.get("plan_date", datetime.today().strftime("%d %b %Y"))),
        ("Client", brief.get("client_name", "")),
        ("Brand", brief.get("brand_name", "")),
        ("Campaign Name", brief.get("campaign_name", "")),
        ("Campaign Start", brief["start_date"].strftime("%d %b %Y") if brief.get("start_date") else ""),
        ("Campaign End",   brief["end_date"].strftime("%d %b %Y")   if brief.get("end_date")   else ""),
        ("Plan Version", "V1.0"),
    ]
    for i, (label, val) in enumerate(meta, 2):
        ws.row_dimensions[i].height = 20
        cs(i, 1, label, bold=True, bg=light, fg=navy, align="right")
        ms(i, 2, i, 5, value=val)
        for c in range(6, 14): cs(i, c, bg=white)

    hdr_row = 10
    ws.row_dimensions[hdr_row].height = 36
    hdrs = ["Channel / Placement","Campaign Objective","Audience","Primary KPI Type","Buying Rate","Primary KPI","Spendable (USD)","Spendable (LKR)","Billable (LKR)","Audience Description","Start Date","End Date","Days"]
    for ci, h in enumerate(hdrs, 1):
        cs(hdr_row, ci, h, bold=True, bg=navy, fg=white, align="center", size=9)

    usd_rate = 320; commission = 0.10; ssc_rate = 0.025641; vat_rate = 0.18; wht_rate = 0.163
    total_lkr = brief.get("total_budget", 0)
    channels  = brief.get("channels", [])
    start_dt  = brief.get("start_date", date.today())
    end_dt    = brief.get("end_date",   date.today())
    days      = (end_dt - start_dt).days

    channel_groups = {
        "Meta":          {"colour": "1877F2", "subs": ["Facebook", "Instagram"]},
        "TikTok":        {"colour": "010101", "subs": ["TikTok In-Feed", "TikTok TopView"]},
        "Google":        {"colour": "1A8A6E", "subs": ["YouTube", "GDN", "Google Search"]},
        "Programmatic":  {"colour": "7B5EA7", "subs": ["Display", "Video"]},
        "LinkedIn":      {"colour": "0A66C2", "subs": ["LinkedIn Feed"]},
        "YouTube":       {"colour": "FF0000", "subs": ["YouTube Pre-roll", "YouTube Bumper"]},
    }

    selected_groups = {}
    for ch in channels:
        for grp in channel_groups:
            if grp.lower() in ch.lower() and grp not in selected_groups:
                selected_groups[grp] = round(total_lkr / len(channels), 0)

    channel_totals = {}
    current_row = hdr_row + 1

    for grp_name, grp_lkr in selected_groups.items():
        grp_info = channel_groups.get(grp_name, {"colour": navy, "subs": [grp_name]})
        ws.row_dimensions[current_row].height = 22
        ms(current_row, 1, current_row, 13, grp_name.upper(), bold=True, bg=grp_info["colour"], fg=white, align="left", size=10)
        current_row += 1
        channel_totals[grp_name] = 0

        subs  = grp_info["subs"]
        sub_n = len(subs)
        for sub in subs:
            sub_lkr = round(grp_lkr / sub_n, 0)
            sub_usd = round(sub_lkr / usd_rate, 2)
            billable = round(sub_lkr * 1.05, 0)
            ws.row_dimensions[current_row].height = 20
            row_bg = "F0F4FA" if current_row % 2 == 0 else white
            data = [sub, brief.get("objective",""), brief.get("target_audience",""), brief.get("kpi_focus",""), "CPM", "", sub_usd, sub_lkr, billable, brief.get("target_audience",""), start_dt.strftime("%d/%m/%Y"), end_dt.strftime("%d/%m/%Y"), days]
            for ci, val in enumerate(data, 1):
                fmt = '#,##0.00' if ci==7 else ('#,##0' if ci in (8,9) else None)
                cs(current_row, ci, val, bg=row_bg, align="center" if ci>5 else "left", num_fmt=fmt)
            channel_totals[grp_name] += sub_lkr
            current_row += 1

    current_row += 1
    total_working = sum(channel_totals.values())
    agency_comm   = round(total_working * commission, 2)
    sub1 = total_working + agency_comm
    ssc  = round(sub1 * ssc_rate, 2)
    sub2 = sub1 + ssc
    vat  = round(sub2 * vat_rate, 2)
    wht  = round(sub2 * wht_rate, 2)
    total_invest = sub2 + vat

    summary = []
    for grp_name, grp_total in channel_totals.items():
        summary.append((f"{grp_name} Investment (incl. fees)", grp_total))
    summary += [
        ("Total Working Investment (LKR)", total_working),
        ("Agency Commission (10%)", agency_comm),
        ("Sub Total", sub1),
        ("SSC Levy (2.5641%)", ssc),
        ("Sub Total", sub2),
        ("VAT (18%)", vat),
        ("Withholding Tax (16.3%)", wht),
        ("TOTAL INVESTMENT (LKR)", total_invest),
    ]
    for label, val in summary:
        ws.row_dimensions[current_row].height = 22
        is_total = "TOTAL INVESTMENT" in label
        is_sub   = label.startswith("Sub Total") or label.startswith("Total Working")
        bg = navy if is_total else (light if is_sub else white)
        fg_col = white if is_total else "1A1D23"
        ms(current_row, 1, current_row, 8, label, bold=is_total or is_sub, bg=bg, fg=fg_col, align="right")
        cs(current_row, 9, val, bold=is_total or is_sub, bg=bg, fg=fg_col, align="right", num_fmt='#,##0.00')
        for c in range(10, 14): cs(current_row, c, bg=white)
        current_row += 1

    ws2 = wb.create_sheet("AI Plan")
    ws2.column_dimensions["A"].width = 120
    ws2.cell(row=1, column=1, value="AI-GENERATED CAMPAIGN PLAN").font = Font(name="Inter", bold=True, size=13, color=navy)
    for i, line in enumerate(plan_text.split("\n"), 3):
        c = ws2.cell(row=i, column=1, value=line)
        c.font = Font(name="Inter", size=10)
        c.alignment = Alignment(wrap_text=True)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ── Session defaults ──────────────────────────────────────────────────────────
for key, default in [("step", 1), ("brief", {}), ("selected_client", None),
                     ("selected_brand", None), ("combined_df", None), ("generated_plan", None)]:
    if key not in st.session_state:
        st.session_state[key] = default

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='padding:8px 0 24px 0;'>
      <div style='font-family:Plus Jakarta Sans,sans-serif;font-size:1.25rem;font-weight:800;color:#1e3a5f;'>🎯 Orchestration</div>
      <div style='font-size:0.75rem;color:#8a93a8;margin-top:2px;'>Digital Planning Platform</div>
    </div>""", unsafe_allow_html=True)

    nav = st.radio("Navigation", ["📊 New Campaign Plan", "📁 Saved Plans", "📈 Data Explorer"],
                   label_visibility="collapsed")
    st.markdown("---")
    st.markdown("<div style='font-size:0.72rem;color:#8a93a8;'>POC Version 1.0 · Phase 1</div>", unsafe_allow_html=True)

# ── Hero ──────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero-header">
  <div class="hero-badge">Digital Planning Platform</div>
  <div class="hero-title">Orchestration-Digital</div>
  <div class="hero-sub">Plan · Analyse · Execute · Report — all in one place</div>
</div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: NEW CAMPAIGN PLAN
# ══════════════════════════════════════════════════════════════════════════════
if nav == "📊 New Campaign Plan":

    sb = get_supabase()

    # ── Step indicator ────────────────────────────────────────────────────────
    steps = ["1 · Campaign Brief", "2 · Historical Data", "3 · Media Plan"]
    cols  = st.columns(3)
    for i, (col, label) in enumerate(zip(cols, steps), 1):
        active  = st.session_state["step"] == i
        done    = st.session_state["step"] > i
        bg      = "#1e3a5f" if active else ("#1a8a6e" if done else "#e8eaf0")
        fg      = "#ffffff" if (active or done) else "#8a93a8"
        icon    = "✓" if done else str(i)
        col.markdown(f"""
        <div style='background:{bg};color:{fg};border-radius:10px;padding:12px 16px;text-align:center;font-weight:600;font-size:0.85rem;'>
          {icon} · {label.split("·")[1].strip()}
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 1: CAMPAIGN BRIEF
    # ══════════════════════════════════════════════════════════════════════════
    if st.session_state["step"] == 1:

        st.markdown('<div class="section-header">Client & Brand</div>', unsafe_allow_html=True)

        clients = get_clients_list(sb)
        client_names = [c["client_name"] for c in clients]

        col_a, col_b = st.columns(2)

        with col_a:
            st.markdown("**Select or create a client**")
            client_choice = st.selectbox("Client", ["— Select existing —", "➕ Add new client"] + client_names,
                                          label_visibility="collapsed")
            if client_choice == "➕ Add new client":
                new_client_name = st.text_input("New client name", placeholder="e.g. Unilever Lanka")
                if st.button("Create Client") and new_client_name:
                    record = create_client_record(sb, new_client_name)
                    if record:
                        st.success(f"Client '{new_client_name}' created!")
                        st.rerun()
            elif client_choice != "— Select existing —":
                selected_client = next((c for c in clients if c["client_name"] == client_choice), None)
                st.session_state["selected_client"] = selected_client

        with col_b:
            if st.session_state.get("selected_client"):
                st.markdown("**Select or create a brand**")
                brands = get_brands_for_client(sb, st.session_state["selected_client"]["id"])
                brand_names = [b["brand_name"] for b in brands]
                brand_choice = st.selectbox("Brand", ["— Select existing —", "➕ Add new brand"] + brand_names,
                                             label_visibility="collapsed")
                if brand_choice == "➕ Add new brand":
                    new_brand_name = st.text_input("New brand name", placeholder="e.g. Sunlight")
                    if st.button("Create Brand") and new_brand_name:
                        record = create_brand_record(sb, st.session_state["selected_client"]["id"], new_brand_name)
                        if record:
                            st.success(f"Brand '{new_brand_name}' created!")
                            st.rerun()
                elif brand_choice != "— Select existing —":
                    selected_brand = next((b for b in brands if b["brand_name"] == brand_choice), None)
                    st.session_state["selected_brand"] = selected_brand
            else:
                st.info("Select a client first to choose a brand.")

        # Campaign details
        if st.session_state.get("selected_client") and st.session_state.get("selected_brand"):
            st.markdown('<div class="section-header">Campaign Details</div>', unsafe_allow_html=True)

            col_c, col_d = st.columns(2)
            with col_c:
                campaign_name = st.text_input("Campaign Name", placeholder="e.g. Q3 2025 Brand Awareness",
                                               value=st.session_state["brief"].get("campaign_name",""))
                objective = st.selectbox("Campaign Objective", [
                    "Brand Awareness","Reach & Frequency","Website Traffic",
                    "Lead Generation","App Installs","E-commerce / Conversions",
                    "Video Views","Engagement"
                ])
                market = st.text_input("Market / Region", value=st.session_state["brief"].get("market","Sri Lanka"))

            with col_d:
                target_audience = st.text_area("Target Audience",
                    placeholder="e.g. 25–45, urban professionals, value seekers",
                    height=100,
                    value=st.session_state["brief"].get("target_audience",""))
                kpi_focus = st.selectbox("Primary KPI", [
                    "ROAS (Return on Ad Spend)","CPA (Cost per Acquisition)",
                    "CPL (Cost per Lead)","CTR (Click-through Rate)",
                    "CPM (Cost per 1000 Impressions)","CPC (Cost per Click)",
                    "Reach","Video Completion Rate"
                ])

            st.markdown('<div class="section-header">Budget & Dates</div>', unsafe_allow_html=True)
            col_e, col_f, col_g = st.columns([2,1,1])
            with col_e:
                total_budget = st.number_input("Total Budget (LKR)", min_value=0, value=st.session_state["brief"].get("total_budget",1000000), step=50000, format="%d")
                st.caption(f"💰 {format_lkr(total_budget)}  ≈  USD {total_budget/320:,.0f}")
            with col_f:
                start_date = st.date_input("Start Date", value=st.session_state["brief"].get("start_date", date.today()))
            with col_g:
                end_date = st.date_input("End Date", value=st.session_state["brief"].get("end_date", date.today()))

            st.markdown('<div class="section-header">Channel Selection</div>', unsafe_allow_html=True)
            cols3 = st.columns(3)
            ch_meta         = cols3[0].checkbox("Meta (Facebook/Instagram)", value=True)
            ch_google       = cols3[1].checkbox("Google Ads", value=True)
            ch_tiktok       = cols3[2].checkbox("TikTok Ads", value=True)
            ch_programmatic = cols3[0].checkbox("Programmatic Display")
            ch_youtube      = cols3[1].checkbox("YouTube")
            ch_linkedin     = cols3[2].checkbox("LinkedIn")

            channels_selected = []
            if ch_meta:         channels_selected.append("Meta (Facebook & Instagram)")
            if ch_google:       channels_selected.append("Google Search & Display")
            if ch_tiktok:       channels_selected.append("TikTok Ads")
            if ch_programmatic: channels_selected.append("Programmatic Display")
            if ch_youtube:      channels_selected.append("YouTube")
            if ch_linkedin:     channels_selected.append("LinkedIn")

            st.markdown("---")
            st.markdown('<div class="green-btn">', unsafe_allow_html=True)
            if st.button("Continue to Historical Data →", use_container_width=True):
                if not campaign_name:
                    st.error("Please enter a campaign name.")
                elif end_date <= start_date:
                    st.error("End date must be after start date.")
                elif not channels_selected:
                    st.error("Please select at least one channel.")
                else:
                    st.session_state["brief"] = {
                        "client_name":    st.session_state["selected_client"]["client_name"],
                        "client_id":      st.session_state["selected_client"]["id"],
                        "brand_name":     st.session_state["selected_brand"]["brand_name"],
                        "brand_id":       st.session_state["selected_brand"]["id"],
                        "campaign_name":  campaign_name,
                        "objective":      objective,
                        "total_budget":   total_budget,
                        "start_date":     start_date,
                        "end_date":       end_date,
                        "target_audience": target_audience,
                        "channels":       channels_selected,
                        "market":         market,
                        "kpi_focus":      kpi_focus,
                        "plan_date":      datetime.today().strftime("%d %b %Y"),
                    }
                    st.session_state["step"] = 2
                    st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 2: HISTORICAL DATA
    # ══════════════════════════════════════════════════════════════════════════
    elif st.session_state["step"] == 2:

        brief      = st.session_state["brief"]
        brand_id   = brief["brand_id"]
        brand_name = brief["brand_name"]
        client_name = brief["client_name"]

        st.markdown(f'<div class="section-header">Historical Data — {client_name} · {brand_name}</div>', unsafe_allow_html=True)

        # Load saved files for this brand
        saved_files = get_brand_data_files(sb, brand_id)

        if saved_files:
            st.success(f"✅ Found {len(saved_files)} saved data file(s) for **{brand_name}** — auto-loaded for planning.")

            col1, col2, col3 = st.columns(3)
            total_rows = sum(f.get("row_count", 0) for f in saved_files)
            with col1: st.markdown(f'<div class="metric-card"><div class="metric-value">{len(saved_files)}</div><div class="metric-label">Saved Files</div></div>', unsafe_allow_html=True)
            with col2: st.markdown(f'<div class="metric-card"><div class="metric-value">{total_rows:,}</div><div class="metric-label">Total Rows</div></div>', unsafe_allow_html=True)
            with col3: st.markdown(f'<div class="metric-card"><div class="metric-value">{brand_name}</div><div class="metric-label">Brand</div></div>', unsafe_allow_html=True)

            # Build combined df from saved data
            dfs = []
            with st.expander("📁 Saved files for this brand"):
                for f in saved_files:
                    col_a, col_b, col_c = st.columns([3, 2, 1])
                    col_a.markdown(f"📄 **{f['file_name']}**")
                    col_b.markdown(f"<span style='color:#8a93a8;font-size:0.82rem;'>{f.get('row_count',0):,} rows · {f['uploaded_at'][:10]}</span>", unsafe_allow_html=True)
                    if col_c.button("🗑", key=f"del_{f['id']}", help="Remove this file"):
                        delete_brand_data_file(sb, f["id"])
                        st.rerun()
                    try:
                        df = pd.read_json(io.StringIO(f["data_json"]))
                        dfs.append(df)
                    except:
                        pass

            if dfs:
                st.session_state["combined_df"] = pd.concat(dfs, ignore_index=True)

        else:
            st.info(f"No historical data saved for **{brand_name}** yet. Upload files below.")

        # Upload new files
        st.markdown('<div class="section-header">Upload New Files</div>', unsafe_allow_html=True)
        st.caption("Upload CSV/Excel exports from Meta, Google, TikTok, or any ad platform. Files are saved under this brand for future use.")

        uploaded_files = st.file_uploader(
            "Drop campaign files here",
            type=["csv","xlsx","xls"],
            accept_multiple_files=True,
            help="Hold Ctrl/Cmd to select multiple files"
        )

        if uploaded_files:
            new_dfs = []
            for f in uploaded_files:
                df = parse_uploaded_file(f)
                if df is not None:
                    new_dfs.append((f.name, df))

            if new_dfs:
                if st.button(f"💾 Save {len(new_dfs)} file(s) to {brand_name}", use_container_width=True):
                    for fname, df in new_dfs:
                        save_brand_data(sb, brand_id, fname, df.to_json(), len(df))
                    st.success(f"✅ {len(new_dfs)} file(s) saved under {brand_name}!")
                    st.rerun()

                # Preview
                preview_df = pd.concat([d for _, d in new_dfs], ignore_index=True)
                with st.expander("Preview uploaded data"):
                    st.dataframe(preview_df.head(20), use_container_width=True)

                # Add to session for immediate use
                if st.session_state.get("combined_df") is not None:
                    st.session_state["combined_df"] = pd.concat(
                        [st.session_state["combined_df"]] + [d for _, d in new_dfs], ignore_index=True)
                else:
                    st.session_state["combined_df"] = pd.concat([d for _, d in new_dfs], ignore_index=True)

        st.markdown("---")
        col_nav1, col_nav2 = st.columns(2)
        with col_nav1:
            if st.button("← Back to Brief", use_container_width=True):
                st.session_state["step"] = 1
                st.rerun()
        with col_nav2:
            st.markdown('<div class="green-btn">', unsafe_allow_html=True)
            if st.button("Generate Media Plan →", use_container_width=True):
                if st.session_state.get("combined_df") is None:
                    st.error("Please upload at least one data file before generating a plan.")
                else:
                    with st.spinner("🤖 Analysing historical data and building your media plan…"):
                        try:
                            ai_client = get_anthropic_client()
                            summary   = summarise_dataframe(st.session_state["combined_df"])
                            plan_text = generate_campaign_plan(ai_client, summary, brief)
                            st.session_state["generated_plan"] = plan_text
                            st.session_state["step"] = 3
                            st.rerun()
                        except Exception as e:
                            st.error(f"Error generating plan: {e}")
            st.markdown('</div>', unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 3: MEDIA PLAN
    # ══════════════════════════════════════════════════════════════════════════
    elif st.session_state["step"] == 3:

        brief = st.session_state["brief"]
        plan_text = st.session_state["generated_plan"]
        duration = (brief["end_date"] - brief["start_date"]).days

        st.markdown(f'<div class="section-header">{brief["client_name"]} · {brief["brand_name"]} · {brief["campaign_name"]}</div>', unsafe_allow_html=True)

        col1, col2, col3, col4 = st.columns(4)
        with col1: st.markdown(f'<div class="metric-card"><div class="metric-value">{format_lkr(brief["total_budget"])}</div><div class="metric-label">Total Budget</div></div>', unsafe_allow_html=True)
        with col2: st.markdown(f'<div class="metric-card"><div class="metric-value">{duration} days</div><div class="metric-label">Duration</div></div>', unsafe_allow_html=True)
        with col3: st.markdown(f'<div class="metric-card"><div class="metric-value">{len(brief["channels"])}</div><div class="metric-label">Channels</div></div>', unsafe_allow_html=True)
        with col4: st.markdown(f'<div class="metric-card"><div class="metric-value">{brief["objective"].split()[0]}</div><div class="metric-label">Objective</div></div>', unsafe_allow_html=True)

        plan_html = plan_text.replace("\n", "<br>")
        st.markdown(f'<div class="plan-card">{plan_html}</div>', unsafe_allow_html=True)

        st.markdown("---")
        col_a, col_b, col_c, col_d = st.columns(4)

        with col_a:
            st.download_button("📥 Download TXT", data=plan_text,
                file_name=f"{brief['campaign_name'].replace(' ','_')}_plan.txt",
                mime="text/plain", use_container_width=True)

        with col_b:
            excel_bytes = build_template_excel(brief, plan_text)
            st.download_button("📊 Download Excel (Template)",
                data=excel_bytes,
                file_name=f"{brief['campaign_name'].replace(' ','_')}_MediaPlan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

        with col_c:
            if st.button("💾 Save to Library", use_container_width=True):
                save_plan(sb, {
                    "brand_id":      brief.get("brand_id"),
                    "client_name":   brief.get("client_name"),
                    "brand_name":    brief.get("brand_name"),
                    "campaign_name": brief.get("campaign_name"),
                    "objective":     brief.get("objective"),
                    "total_budget":  brief.get("total_budget"),
                    "start_date":    str(brief.get("start_date")),
                    "end_date":      str(brief.get("end_date")),
                    "channels":      json.dumps(brief.get("channels",[])),
                    "market":        brief.get("market"),
                    "kpi_focus":     brief.get("kpi_focus"),
                    "plan_text":     plan_text,
                    "created_at":    datetime.utcnow().isoformat()
                })
                st.success("Saved to library!")

        with col_d:
            st.markdown('<div class="green-btn">', unsafe_allow_html=True)
            if st.button("➕ New Plan", use_container_width=True):
                st.session_state["step"] = 1
                st.session_state["generated_plan"] = None
                st.session_state["brief"] = {}
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: SAVED PLANS
# ══════════════════════════════════════════════════════════════════════════════
elif nav == "📁 Saved Plans":
    sb = get_supabase()
    st.markdown('<div class="section-header">Saved Campaign Plans</div>', unsafe_allow_html=True)
    try:
        plans = load_saved_plans(sb)
        if not plans:
            st.info("No saved plans yet.")
        else:
            for plan in plans:
                label = f"📋 {plan.get('client_name','—')} · {plan.get('brand_name','—')} · {plan.get('campaign_name','Unnamed')} — {format_lkr(plan.get('total_budget',0))}"
                with st.expander(label):
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
    if st.session_state.get("combined_df") is None:
        st.info("Complete Step 2 in New Campaign Plan to load data here.")
    else:
        df = st.session_state["combined_df"]
        st.caption(f"{len(df):,} rows · {len(df.columns)} columns")
        st.dataframe(df, use_container_width=True)
        numeric_cols = df.select_dtypes(include='number').columns.tolist()
        if len(numeric_cols) >= 2:
            st.markdown('<div class="section-header">Quick Chart</div>', unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            x_col = c1.selectbox("X axis", df.columns.tolist())
            y_col = c2.selectbox("Y axis", numeric_cols)
            chart_type = st.radio("Chart type", ["Line","Bar","Area"], horizontal=True)
            chart_df = df[[x_col, y_col]].dropna().set_index(x_col)
            if chart_type == "Line": st.line_chart(chart_df)
            elif chart_type == "Bar": st.bar_chart(chart_df)
            else: st.area_chart(chart_df)