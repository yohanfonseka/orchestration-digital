import streamlit as st
import pandas as pd
import numpy as np
import json
import anthropic
from datetime import datetime, date
import io
from supabase import create_client, Client
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Orchestration-Digital", page_icon="🎯", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
  html, body, [class*="css"] { font-family:'Inter',sans-serif; background-color:#ffffff !important; color:#1a1d23; }
  .stApp { background:#f8f9fc !important; }

  /* ── Sidebar ── */
  [data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1e3a5f 0%, #162d4a 100%) !important;
    border-right: none !important;
  }
  [data-testid="stSidebar"] * { color: #ffffff !important; }
  [data-testid="stSidebar"] .stRadio label {
    display: block !important;
    padding: 10px 16px !important;
    border-radius: 8px !important;
    font-size: 0.88rem !important;
    font-weight: 500 !important;
    color: rgba(255,255,255,0.75) !important;
    cursor: pointer !important;
    transition: all 0.15s !important;
    margin-bottom: 2px !important;
  }
  [data-testid="stSidebar"] .stRadio label:hover {
    background: rgba(255,255,255,0.1) !important;
    color: #ffffff !important;
  }
  [data-testid="stSidebar"] .stRadio [aria-checked="true"] + label,
  [data-testid="stSidebar"] .stRadio label:has(input:checked) {
    background: rgba(255,255,255,0.15) !important;
    color: #ffffff !important;
    font-weight: 600 !important;
  }
  [data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.15) !important; }
  [data-testid="stSidebar"] .stRadio > div { gap: 0 !important; }

  /* ── Hero ── */
  .hero-header { background:linear-gradient(135deg,#1e3a5f 0%,#2d5a9b 50%,#1a8a6e 100%); border-radius:16px; padding:36px 40px; margin-bottom:28px; }
  .hero-badge { display:inline-block; background:rgba(255,255,255,0.15); color:#fff; border:1px solid rgba(255,255,255,0.25); border-radius:20px; font-size:0.7rem; font-weight:600; letter-spacing:0.1em; text-transform:uppercase; padding:4px 14px; margin-bottom:12px; }
  .hero-title { font-family:'Plus Jakarta Sans',sans-serif; font-size:2rem; font-weight:800; color:#fff; letter-spacing:-0.5px; margin:0 0 8px 0; }
  .hero-sub { font-size:0.9rem; color:rgba(255,255,255,0.75); margin:0; }

  /* ── Cards ── */
  .metric-card { background:#fff; border:1px solid #e8eaf0; border-radius:12px; padding:20px 24px; text-align:center; box-shadow:0 1px 4px rgba(0,0,0,0.05); }
  .metric-value { font-family:'Plus Jakarta Sans',sans-serif; font-size:1.6rem; font-weight:700; color:#1e3a5f; }
  .metric-label { font-size:0.75rem; color:#8a93a8; text-transform:uppercase; letter-spacing:0.06em; margin-top:4px; }
  .section-header { font-family:'Plus Jakarta Sans',sans-serif; font-size:1rem; font-weight:700; color:#1a1d23; border-left:3px solid #2d5a9b; padding-left:12px; margin:24px 0 14px 0; }
  .settings-card { background:#fff; border:1px solid #e8eaf0; border-radius:12px; padding:20px 24px; margin-bottom:16px; box-shadow:0 1px 4px rgba(0,0,0,0.05); }

  /* ── Chat ── */
  .chat-container { background:#fff; border:1px solid #e8eaf0; border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.05); margin-bottom:16px; }
  .chat-header { background:linear-gradient(135deg,#1e3a5f,#2d5a9b); padding:16px 20px; }
  .chat-header-title { color:#fff; font-family:'Plus Jakarta Sans',sans-serif; font-weight:700; font-size:0.95rem; margin:0; }
  .chat-header-sub { color:rgba(255,255,255,0.7); font-size:0.78rem; margin:2px 0 0 0; }
  .chat-messages { padding:20px; max-height:520px; overflow-y:auto; scroll-behavior:smooth; }
  .msg-agent { display:flex; gap:12px; margin-bottom:16px; align-items:flex-start; }
  .msg-user { display:flex; gap:12px; margin-bottom:16px; align-items:flex-start; flex-direction:row-reverse; }
  .avatar-agent { width:32px; height:32px; background:linear-gradient(135deg,#1e3a5f,#2d5a9b); border-radius:50%; display:flex; align-items:center; justify-content:center; color:#fff; font-size:0.75rem; font-weight:700; flex-shrink:0; }
  .avatar-user { width:32px; height:32px; background:linear-gradient(135deg,#1a8a6e,#22b894); border-radius:50%; display:flex; align-items:center; justify-content:center; color:#fff; font-size:0.75rem; font-weight:700; flex-shrink:0; }
  .bubble-agent { background:#f0f4fa; border-radius:0 12px 12px 12px; padding:12px 16px; max-width:80%; font-size:0.88rem; line-height:1.6; color:#1a1d23; }
  .bubble-user { background:linear-gradient(135deg,#1e3a5f,#2d5a9b); border-radius:12px 0 12px 12px; padding:12px 16px; max-width:80%; font-size:0.88rem; line-height:1.6; color:#fff; }
  #chat-bottom { height:1px; }

  /* ── Buttons ── */
  .stButton > button { background:linear-gradient(135deg,#1e3a5f,#2d5a9b) !important; color:white !important; border:none !important; border-radius:8px !important; font-weight:600 !important; font-size:0.9rem !important; padding:10px 24px !important; }
  .stButton > button:hover { opacity:0.88 !important; }
  .green-btn .stButton > button { background:linear-gradient(135deg,#1a8a6e,#22b894) !important; }
  .plan-card { background:#fff; border:1px solid #e8eaf0; border-radius:14px; padding:28px 32px; margin-top:16px; line-height:1.8; color:#2d3341; box-shadow:0 1px 4px rgba(0,0,0,0.05); font-size:0.92rem; }

  /* ── Inputs ── */
  .stTextInput > div > div > input, .stNumberInput > div > div > input,
  .stTextArea textarea, .stDateInput > div > div > input {
    background:#fff !important; border:1px solid #d0d5e0 !important; border-radius:8px !important; color:#1a1d23 !important; font-size:0.9rem !important;
  }
  .stSelectbox > div > div { background:#fff !important; border:1px solid #d0d5e0 !important; border-radius:8px !important; color:#1a1d23 !important; }
  .stTabs [data-baseweb="tab-list"] { background:#eef1f8; border-radius:10px; padding:4px; gap:4px; }
  .stTabs [data-baseweb="tab"] { background:transparent !important; color:#6b7590 !important; border-radius:7px !important; font-weight:500 !important; font-size:0.88rem !important; }
  .stTabs [aria-selected="true"] { background:#fff !important; color:#1e3a5f !important; font-weight:600 !important; box-shadow:0 1px 4px rgba(0,0,0,0.08) !important; }
  div[data-testid="stExpander"] { background:#fff !important; border:1px solid #e8eaf0 !important; border-radius:10px !important; }
  label { color:#4a5168 !important; font-size:0.85rem !important; font-weight:500 !important; }
  #MainMenu, footer, header { visibility:hidden; }
</style>
""", unsafe_allow_html=True)

# Auto-scroll JS
AUTO_SCROLL_JS = """
<script>
  function scrollChat() {
    const el = document.getElementById('chat-bottom');
    if (el) el.scrollIntoView({ behavior: 'smooth' });
    const msgs = document.querySelector('.chat-messages');
    if (msgs) msgs.scrollTop = msgs.scrollHeight;
  }
  setTimeout(scrollChat, 120);
  setTimeout(scrollChat, 400);
</script>
"""

@st.cache_resource
def get_supabase():
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

def get_anthropic_client():
    return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])

# ── Supabase helpers ──────────────────────────────────────────────────────────
def get_clients_list(sb):
    try: return sb.table("clients").select("*").order("client_name").execute().data or []
    except: return []

def get_brands_for_client(sb, client_id):
    try: return sb.table("brands").select("*").eq("client_id", client_id).order("brand_name").execute().data or []
    except: return []

def create_client_record(sb, name):
    try:
        r = sb.table("clients").insert({"client_name": name, "created_at": datetime.utcnow().isoformat()}).execute()
        return r.data[0] if r.data else None
    except Exception as e: st.error(f"Error: {e}"); return None

def create_brand_record(sb, client_id, name):
    try:
        r = sb.table("brands").insert({"client_id": client_id, "brand_name": name, "created_at": datetime.utcnow().isoformat()}).execute()
        return r.data[0] if r.data else None
    except Exception as e: st.error(f"Error: {e}"); return None

def get_brand_data_files(sb, brand_id):
    try: return sb.table("brand_data").select("*").eq("brand_id", brand_id).order("uploaded_at", desc=True).execute().data or []
    except: return []

def save_brand_data(sb, brand_id, file_name, data_json, row_count):
    try: sb.table("brand_data").insert({"brand_id": brand_id, "file_name": file_name, "data_json": data_json, "row_count": row_count, "uploaded_at": datetime.utcnow().isoformat()}).execute(); return True
    except Exception as e: st.error(f"Error: {e}"); return False

def delete_brand_data_file(sb, file_id):
    try: sb.table("brand_data").delete().eq("id", file_id).execute(); return True
    except: return False

def save_plan(sb, plan_data):
    try: sb.table("campaign_plans").insert(plan_data).execute(); return True
    except Exception as e: st.warning(f"Could not save: {e}"); return False

def load_saved_plans(sb):
    try: return sb.table("campaign_plans").select("*").order("created_at", desc=True).limit(50).execute().data or []
    except: return []

def get_plan_version(sb, campaign_name, brand_id):
    try:
        plans = sb.table("campaign_plans").select("plan_version").eq("campaign_name", campaign_name).eq("brand_id", str(brand_id)).execute().data or []
        if not plans: return "V1.0"
        versions = []
        for p in plans:
            try: versions.append(float(p.get("plan_version","V1.0").replace("V","")))
            except: versions.append(1.0)
        return f"V{max(versions)+1:.1f}"
    except: return "V1.0"

def get_platform_settings(sb):
    try: return sb.table("platform_settings").select("*").order("platform_name").execute().data or []
    except: return []

def save_platform_setting(sb, platform_name, data):
    try:
        existing = sb.table("platform_settings").select("id").eq("platform_name", platform_name).execute().data
        if existing: sb.table("platform_settings").update({**data, "updated_at": datetime.utcnow().isoformat()}).eq("platform_name", platform_name).execute()
        else: sb.table("platform_settings").insert({"platform_name": platform_name, **data, "updated_at": datetime.utcnow().isoformat()}).execute()
        return True
    except Exception as e: st.error(f"Error: {e}"); return False

# ── Data helpers ──────────────────────────────────────────────────────────────
def parse_uploaded_file(f):
    n = f.name.lower()
    if n.endswith(".csv"): return pd.read_csv(f)
    elif n.endswith((".xlsx",".xls")): return pd.read_excel(f)
    return None

def extract_channel_benchmarks(df):
    benchmarks = {}
    col_lower = {c.lower(): c for c in df.columns}
    platform_keywords = {
        "Facebook":["facebook","fb"],"Instagram":["instagram","ig"],
        "TikTok":["tiktok","tik tok"],"YouTube":["youtube","yt"],
        "Google Search":["google search","search","sem"],
        "Google Display":["google display","display","gdn"],
        "LinkedIn":["linkedin"],"Programmatic":["programmatic","dsp"],
    }
    channel_col = next((c for cl,c in col_lower.items() if any(k in cl for k in ["channel","platform","campaign","source"])), None)

    def get_avg(sub_df, key):
        matches = [c for cl,c in {cc.lower():cc for cc in sub_df.columns}.items() if key in cl]
        if matches:
            try:
                vals = pd.to_numeric(sub_df[matches[0]], errors="coerce").dropna()
                return round(float(vals.mean()),2) if len(vals)>0 else None
            except: return None
        return None

    if channel_col:
        for platform, keywords in platform_keywords.items():
            mask = df[channel_col].astype(str).str.lower().apply(lambda x: any(kw in x for kw in keywords))
            sub = df[mask]
            if len(sub)>0:
                benchmarks[platform] = {
                    "cpm":get_avg(sub,"cpm"),"cpc":get_avg(sub,"cpc"),
                    "cpv":get_avg(sub,"cpv"),"cpa":get_avg(sub,"cpa"),
                    "ctr":get_avg(sub,"ctr"),"roas":get_avg(sub,"roas"),"rows":len(sub)
                }
    # Global fallback
    global_stats = {
        "cpm":get_avg(df,"cpm"),"cpc":get_avg(df,"cpc"),
        "cpv":get_avg(df,"cpv"),"cpa":get_avg(df,"cpa"),
        "ctr":get_avg(df,"ctr"),"roas":get_avg(df,"roas"),"rows":len(df)
    }
    for platform in platform_keywords:
        if platform not in benchmarks:
            benchmarks[platform] = {**global_stats, "is_global_fallback": True}
    return benchmarks

# Sri Lanka industry averages (LKR)
INDUSTRY_AVERAGES = {
    "Facebook":      {"cpm":420,"cpc":85,"cpv":None,"cpa":950,"ctr":1.2,"roas":2.8},
    "Instagram":     {"cpm":480,"cpc":95,"cpv":None,"cpa":1100,"ctr":1.0,"roas":2.5},
    "YouTube":       {"cpm":350,"cpc":None,"cpv":9,"cpa":None,"ctr":0.4,"roas":None},
    "Google Search": {"cpm":None,"cpc":75,"cpv":None,"cpa":820,"ctr":3.5,"roas":4.2},
    "Google Display":{"cpm":280,"cpc":60,"cpv":None,"cpa":1200,"ctr":0.35,"roas":None},
    "TikTok":        {"cpm":380,"cpc":90,"cpv":7,"cpa":1050,"ctr":1.5,"roas":2.2},
    "LinkedIn":      {"cpm":1200,"cpc":320,"cpv":None,"cpa":2500,"ctr":0.6,"roas":None},
    "Programmatic":  {"cpm":250,"cpc":55,"cpv":None,"cpa":1400,"ctr":0.25,"roas":None},
}

def calculate_budget_split(channels, total_budget, objective, audience_sizes, benchmarks):
    objective_weights = {
        "Brand Awareness":         {"Facebook":1.3,"Instagram":1.2,"YouTube":1.4,"Google Display":1.0,"TikTok":1.1,"Google Search":0.6,"LinkedIn":0.8,"Programmatic":0.9},
        "Reach & Frequency":       {"Facebook":1.4,"Instagram":1.3,"YouTube":1.3,"Google Display":0.9,"TikTok":1.0,"Google Search":0.5,"LinkedIn":0.7,"Programmatic":0.8},
        "Video Views":             {"YouTube":1.5,"TikTok":1.4,"Facebook":1.1,"Instagram":1.2,"Google Display":0.8,"Google Search":0.4,"LinkedIn":0.7,"Programmatic":0.9},
        "Website Traffic":         {"Google Search":1.5,"Facebook":1.1,"Instagram":1.0,"TikTok":0.9,"Google Display":1.0,"YouTube":0.8,"LinkedIn":1.0,"Programmatic":0.9},
        "Lead Generation":         {"Google Search":1.4,"Facebook":1.3,"LinkedIn":1.5,"Instagram":1.1,"TikTok":0.8,"YouTube":0.7,"Google Display":0.9,"Programmatic":0.8},
        "E-commerce / Conversions":{"Google Search":1.5,"Facebook":1.3,"Instagram":1.2,"TikTok":1.0,"YouTube":0.9,"Google Display":1.0,"LinkedIn":0.7,"Programmatic":1.1},
        "App Installs":            {"Facebook":1.3,"Instagram":1.2,"TikTok":1.4,"Google Search":1.1,"YouTube":1.0,"Google Display":0.9,"LinkedIn":0.7,"Programmatic":0.9},
        "Engagement":              {"Instagram":1.4,"TikTok":1.5,"Facebook":1.2,"YouTube":1.0,"LinkedIn":0.9,"Google Search":0.5,"Google Display":0.7,"Programmatic":0.7},
    }
    obj_key = next((k for k in objective_weights if k.lower() in objective.lower()), "Brand Awareness")
    weights = objective_weights[obj_key]
    scores = {}
    for ch in channels:
        wt = next((v for k,v in weights.items() if k.lower() in ch.lower()), 1.0)
        aud = audience_sizes.get(ch, 500000)
        aud_score = np.log10(max(aud,1000)) / np.log10(10_000_000)
        bench = benchmarks.get(ch, {})
        cpm = bench.get("cpm") or INDUSTRY_AVERAGES.get(ch,{}).get("cpm",450)
        efficiency_score = min(1000/max(cpm,50), 2.0)
        scores[ch] = wt * aud_score * efficiency_score
    total_score = sum(scores.values()) or 1
    split = {ch: round((scores[ch]/total_score)*total_budget,0) for ch in channels}
    min_budget = total_budget * 0.05
    for ch in split:
        if split[ch] < min_budget: split[ch] = min_budget
    total_alloc = sum(split.values())
    return {ch: round(v/total_alloc*total_budget,0) for ch,v in split.items()}

def calculate_channel_kpis(channel, budget_lkr, benchmarks, objective):
    bench = benchmarks.get(channel, {})
    industry = INDUSTRY_AVERAGES.get(channel, {})
    objective_kpi_map = {
        "awareness":("CPM","Impressions"),"reach":("CPM","Impressions"),
        "video":("CPV","Video Views"),"traffic":("CPC","Clicks"),
        "lead":("CPL","Leads"),"conversion":("CPA","Conversions"),
        "ecommerce":("CPA","Conversions"),"app":("CPC","Installs"),
        "engagement":("CPE","Engagements"),
    }
    obj_lower = objective.lower()
    kpi_type, kpi_metric = next(((v[0],v[1]) for k,v in objective_kpi_map.items() if k in obj_lower),("CPM","Impressions"))
    if "search" in channel.lower(): kpi_type,kpi_metric = "CPC","Clicks"
    elif "youtube" in channel.lower(): kpi_type,kpi_metric = "CPV","Video Views"
    elif "linkedin" in channel.lower() and "lead" in obj_lower: kpi_type,kpi_metric = "CPL","Leads"

    rate_key = kpi_type.lower().replace("cpl","cpa").replace("cpe","cpc")
    historical_rate = bench.get(rate_key)
    is_industry_avg = False
    data_source = "Historical Data"

    if not historical_rate or bench.get("is_global_fallback"):
        industry_rate = industry.get(rate_key)
        if industry_rate:
            historical_rate = industry_rate
            is_industry_avg = True
            data_source = "⚠️ Industry Average (LK)"
        else:
            fallback = {"cpm":450,"cpc":95,"cpv":12,"cpa":850}
            historical_rate = fallback.get(rate_key,450)
            is_industry_avg = True
            data_source = "⚠️ Industry Average (LK)"

    buying_rate_lkr = round(historical_rate, 2)
    if kpi_type=="CPM":
        target=int(budget_lkr/buying_rate_lkr*1000); target_str=f"{target:,} Impressions"; rate_str=f"LKR {buying_rate_lkr:,.0f} CPM"
    elif kpi_type=="CPC":
        target=int(budget_lkr/buying_rate_lkr); target_str=f"{target:,} Clicks"; rate_str=f"LKR {buying_rate_lkr:,.0f} CPC"
    elif kpi_type=="CPV":
        target=int(budget_lkr/buying_rate_lkr); target_str=f"{target:,} Video Views"; rate_str=f"LKR {buying_rate_lkr:,.2f} CPV"
    elif kpi_type in ("CPA","CPL"):
        target=int(budget_lkr/buying_rate_lkr); target_str=f"{target:,} {kpi_metric}"; rate_str=f"LKR {buying_rate_lkr:,.0f} {kpi_type}"
    elif kpi_type=="ROAS":
        roas=bench.get("roas") or industry.get("roas") or 3.0; target_str=f"{roas:.1f}x ROAS"; rate_str=f"{roas:.1f}x"
    else:
        target=int(budget_lkr/buying_rate_lkr*1000); target_str=f"{target:,} Impressions"; rate_str=f"LKR {buying_rate_lkr:,.0f}"

    ctr = bench.get("ctr") or (industry.get("ctr") if is_industry_avg else None)
    return {"kpi_type":kpi_type,"buying_rate":rate_str,"target_kpi":target_str,
            "ctr_bench":f"CTR: {ctr:.2f}%" if ctr else "","raw_rate":buying_rate_lkr,
            "is_industry_avg":is_industry_avg,"data_source":data_source}

def get_data_gaps(channels, benchmarks):
    gaps = []
    for ch in channels:
        bench = benchmarks.get(ch,{})
        if not bench or bench.get("is_global_fallback") or (not bench.get("cpm") and not bench.get("cpc")):
            ia = INDUSTRY_AVERAGES.get(ch,{})
            gaps.append(f"**{ch}**: No historical data found — using Sri Lanka industry average "
                       f"(CPM: LKR {ia.get('cpm','—')}, CPC: LKR {ia.get('cpc','—')})")
    return gaps

def summarise_dataframe(df):
    """Compact summary — benchmarks only, no raw data rows. Saves ~60% tokens."""
    col_lower = {c.lower():c for c in df.columns}
    lines = [f"Dataset: {len(df):,} rows, {len(df.columns)} columns",
             f"Columns: {', '.join(df.columns.tolist())}","","Performance benchmarks:"]
    for metric in ["cpm","cpc","ctr","roas","cpa","cpv","spend","impressions","clicks","conversions","reach","frequency"]:
        matches = [c for cl,c in col_lower.items() if metric in cl]
        if matches:
            try:
                vals = pd.to_numeric(df[matches[0]],errors="coerce").dropna()
                if len(vals)>0:
                    lines.append(f"  {matches[0]}: avg={vals.mean():.2f}, median={vals.median():.2f}, min={vals.min():.2f}, max={vals.max():.2f} (n={len(vals)})")
            except: pass
    # Channel breakdown if available
    ch_col = next((c for cl,c in col_lower.items() if any(k in cl for k in ["channel","platform","campaign name","source"])), None)
    if ch_col:
        lines += ["","Channel breakdown:"]
        for ch_val in df[ch_col].dropna().unique()[:12]:
            sub = df[df[ch_col]==ch_val]
            sub_metrics = []
            for m in ["cpm","cpc","ctr","roas"]:
                sub_matches = [c for cl,c in {cc.lower():cc for cc in sub.columns}.items() if m in cl]
                if sub_matches:
                    try:
                        v = pd.to_numeric(sub[sub_matches[0]],errors="coerce").dropna()
                        if len(v)>0: sub_metrics.append(f"{m}={v.mean():.2f}")
                    except: pass
            if sub_metrics:
                lines.append(f"  {ch_val} ({len(sub)} rows): {', '.join(sub_metrics)}")
    return "\n".join(lines)

def format_lkr(v):
    try: return f"LKR {float(v):,.0f}"
    except: return "LKR 0"

# ── Chat ──────────────────────────────────────────────────────────────────────
def render_message(role, content):
    avatar = "AI" if role=="agent" else "You"
    avatar_cls = "avatar-agent" if role=="agent" else "avatar-user"
    bubble_cls = "bubble-agent" if role=="agent" else "bubble-user"
    msg_cls = "msg-agent" if role=="agent" else "msg-user"
    st.markdown(f"""<div class="{msg_cls}"><div class="{avatar_cls}">{avatar}</div><div class="{bubble_cls}">{content.replace(chr(10),'<br>')}</div></div>""", unsafe_allow_html=True)

PLANNING_SYSTEM = """You are Orchy, an expert digital media planning agent inside Orchestration-Digital, used by planners in Sri Lanka.

Collect everything needed for a complete digital media plan through friendly, professional conversation.

CONVERSATION FLOW — cover in order, 1–2 questions at a time:
1. Campaign name and objective (Awareness / Traffic / Leads / Conversions / App Installs / Engagement / Video Views)
2. Total budget in LKR and flight dates
3. Target audience description (demographics, interests, behaviours)
4. Channels — ask about each: Facebook, Instagram, YouTube, Google Search, Google Display, TikTok, LinkedIn, Programmatic Display
5. For each selected channel, ask for the TARGETABLE AUDIENCE SIZE from the platform estimator (Meta Audience Insights, Google Reach Planner, TikTok Audience Estimator). Be specific about where to find this.
6. For each channel, ask what creative assets are available: Static Images, Videos, Carousels, Stories, UGC, Influencer Content
   - For each asset type, confirm: placement, objective, and KPI
   - Ask: "Would you like to allocate separate budgets per creative type, or keep the total budget at platform level?"
   - If per-creative: ask for the budget split across creative types for that channel
7. Any additional context or constraints

RULES:
- Conversational and professional, 1–2 questions at a time
- Confirm back what you've heard
- When ALL info collected, end with exactly: [BRIEF_COMPLETE]
- Never generate the plan yourself
- All budgets in LKR
- When asking for audience sizes, explain exactly where to find them on each platform"""

EDIT_SYSTEM = """You are Orchy, an expert digital media planning agent. You are helping a planner edit an existing campaign plan.

The original plan is provided below. Your job is to:
1. Understand what changes the planner wants to make
2. Ask clarifying questions if needed
3. When you have enough information, confirm the changes and end with exactly: [EDIT_COMPLETE]

Be specific — confirm budget numbers, channel changes, and KPI impacts before finalising.
All budgets in LKR."""

def get_agent_response(messages, client_name, brand_name, data_summary, mode="planning", existing_plan=""):
    client = get_anthropic_client()
    if mode=="editing":
        system = EDIT_SYSTEM + f"\n\nCLIENT: {client_name}\nBRAND: {brand_name}\n\nORIGINAL PLAN:\n{existing_plan[:2000]}"
    else:
        # Data summary only in system prompt once — not repeated in messages
        system = PLANNING_SYSTEM + f"\n\nCLIENT: {client_name}\nBRAND: {brand_name}\n\nHISTORICAL BENCHMARKS:\n{data_summary[:1500]}"
    # Keep only last 8 messages to limit context tokens
    trimmed = messages[-8:] if len(messages) > 8 else messages
    # If trimmed, prepend a brief context note
    if len(messages) > 8:
        earlier = messages[:-8]
        summary_note = f"[Earlier in conversation: {len(earlier)} messages covering initial brief details]"
        trimmed = [{"role":"user","content":summary_note}] + trimmed
    api_messages = [{"role":m["role"].replace("agent","assistant"),"content":m["content"]} for m in trimmed]
    response = client.messages.create(model="claude-sonnet-4-6",max_tokens=800,system=system,messages=api_messages)
    return response.content[0].text

def extract_brief_from_conversation(conversation):
    """Extract brief JSON from conversation — uses last 12 messages only to save tokens."""
    client = get_anthropic_client()
    # Use last 12 messages (where all key info will be)
    trimmed = conversation[-12:] if len(conversation) > 12 else conversation
    conv_text = "\n".join([f"{'PLANNER' if m['role']=='user' else 'AGENT'}: {m['content']}" for m in trimmed])
    prompt = f"""Extract the campaign brief. Return ONLY valid JSON with this exact structure:
{{
  "campaign_name": "",
  "objective": "",
  "total_budget": 0,
  "start_date": "YYYY-MM-DD",
  "end_date": "YYYY-MM-DD",
  "audience": "",
  "channels": [],
  "audience_sizes": {{}},
  "assets": "",
  "market": "Sri Lanka",
  "channel_placements": {{
    "ChannelName": [
      {{"placement": "Feed", "kpi_type": "CPM", "assets": "Static Image", "objective": "Awareness", "budget": 0}}
    ]
  }}
}}

Notes:
- channel_placements: only populate if the planner specified per-creative budget splits. If platform-level only, leave as empty object {{}}.
- budget in channel_placements is in LKR. Sum of placements for a channel should equal that channel's total budget.
- If no per-creative split was discussed, leave channel_placements as {{}}.

CONVERSATION:
{conv_text}

JSON only, no markdown."""
    msg = client.messages.create(model="claude-sonnet-4-6",max_tokens=600,messages=[{"role":"user","content":prompt}])
    try:
        text=msg.content[0].text.strip().replace("```json","").replace("```","").strip()
        return json.loads(text)
    except:
        return {"campaign_name":"Campaign","objective":"","total_budget":0,"start_date":str(date.today()),
                "end_date":str(date.today()),"audience":"","channels":[],"audience_sizes":{},"assets":"","market":"Sri Lanka"}

def generate_media_plan(conversation, client_name, brand_name, data_summary, budget_split, channel_kpi_data, data_gaps):
    """Optimised: uses pre-calculated numbers + compact conversation summary instead of raw history."""
    client = get_anthropic_client()
    # Use only last 10 conversation messages — key details already extracted into budget_split/channel_kpi_data
    trimmed = conversation[-10:] if len(conversation) > 10 else conversation
    conv_text = "\n".join([f"{'PLANNER' if m['role']=='user' else 'AGENT'}: {m['content']}" for m in trimmed])
    split_lines = [f"  {ch}: LKR {b:,.0f} | {channel_kpi_data.get(ch,{}).get('kpi_type','CPM')} | Buying Rate: {channel_kpi_data.get(ch,{}).get('buying_rate','—')} | Target: {channel_kpi_data.get(ch,{}).get('target_kpi','—')} | Source: {channel_kpi_data.get(ch,{}).get('data_source','—')}" for ch,b in budget_split.items()]
    gaps_note = ("\nINDUSTRY AVERAGES USED FOR: " + ", ".join([g.split(":")[0].replace("**","") for g in data_gaps])) if data_gaps else ""

    prompt = f"""Senior digital media planner. Generate a concise, structured media plan.

CLIENT: {client_name} | BRAND: {brand_name}

BUDGET & KPIs (pre-calculated — use exactly):
{chr(10).join(split_lines)}{gaps_note}

KEY BRIEF DETAILS:
{conv_text}

HISTORICAL BENCHMARKS:
{data_summary[:800]}

OUTPUT FORMAT — keep each section tight and use tables where possible:

## 1. EXECUTIVE SUMMARY
2-3 sentences only.

## 2. CHANNEL STRATEGY
One paragraph explaining budget split rationale.

## 3. CHANNEL PLAN
For each channel, use this exact table format:
| Metric | Value |
|--------|-------|
| Budget | LKR X |
| KPI Type | X |
| Buying Rate | X |
| Target KPI | X |
| Ad Formats | X |
| Targeting | X |
| Creative Assets | X |
| Data Source | X |

## 4. CREATIVE ASSET MATRIX
| Asset Type | Platform | Placement | Objective | KPI Target |
|------------|----------|-----------|-----------|------------|

## 5. KPI SUMMARY
| Channel | Budget (LKR) | KPI Type | Buying Rate | Target KPI |
|---------|-------------|----------|-------------|------------|

## 6. OPTIMISATION ROADMAP
| Week | Actions |
|------|---------|

## 7. RISK FLAGS
Bullet points only.

Be precise. Use LKR. Flag ⚠️ where industry averages used."""

    msg = client.messages.create(model="claude-sonnet-4-6",max_tokens=3500,messages=[{"role":"user","content":prompt}])
    return msg.content[0].text

def apply_plan_edits(edit_conversation, original_plan, client_name, brand_name, budget_split, channel_kpi_data):
    client = get_anthropic_client()
    conv_text = "\n".join([f"{'PLANNER' if m['role']=='user' else 'AGENT'}: {m['content']}" for m in edit_conversation])
    split_lines = [f"  {ch}: LKR {b:,.0f} | {channel_kpi_data.get(ch,{}).get('kpi_type','CPM')} | Rate: {channel_kpi_data.get(ch,{}).get('buying_rate','—')} | Target: {channel_kpi_data.get(ch,{}).get('target_kpi','—')}" for ch,b in budget_split.items()]
    prompt = f"""You are a senior digital media planner. Apply the requested edits to this media plan.

CLIENT: {client_name} | BRAND: {brand_name}

ORIGINAL BUDGET SPLIT:
{chr(10).join(split_lines)}

EDIT INSTRUCTIONS FROM PLANNER:
{conv_text}

ORIGINAL PLAN:
{original_plan}

Produce the complete updated media plan with all edits applied. Keep the same structure but update all affected numbers, budgets, KPIs, and rationale. Clearly note what was changed at the top under "CHANGES FROM PREVIOUS VERSION"."""

    msg = client.messages.create(model="claude-sonnet-4-6",max_tokens=5000,messages=[{"role":"user","content":prompt}])
    return msg.content[0].text

# ── Excel ──────────────────────────────────────────────────────────────────────
def build_excel(brief_summary, plan_text="", client_name="", brand_name="", budget_split={}, channel_kpi_data={}, plan_version="V1.0"):
    wb = Workbook(); ws = wb.active; ws.title = "Media Plan"
    navy="1E3A5F"; white="FFFFFF"; light="F0F4FA"; border_col="D0D5E0"

    def cs(row,col,value="",bold=False,bg=None,fg="1A1D23",align="left",size=10,num_fmt=None):
        c=ws.cell(row=row,column=col,value=value)
        c.font=Font(name="Inter",bold=bold,color=fg,size=size)
        if bg: c.fill=PatternFill("solid",fgColor=bg)
        c.alignment=Alignment(horizontal=align,vertical="center",wrap_text=True)
        thin=Side(style="thin",color=border_col)
        c.border=Border(left=thin,right=thin,top=thin,bottom=thin)
        if num_fmt: c.number_format=num_fmt
        return c

    def ms(r1,c1,r2,c2,value="",bold=False,bg=None,fg="1A1D23",align="left",size=10):
        ws.merge_cells(start_row=r1,start_column=c1,end_row=r2,end_column=c2)
        c=ws.cell(row=r1,column=c1,value=value)
        c.font=Font(name="Inter",bold=bold,color=fg,size=size)
        if bg: c.fill=PatternFill("solid",fgColor=bg)
        c.alignment=Alignment(horizontal=align,vertical="center",wrap_text=True)
        return c

    for i,w in enumerate([24,30,20,14,18,20,16,18,18,22,12,12,10],1):
        ws.column_dimensions[get_column_letter(i)].width=w

    ws.row_dimensions[1].height=40
    ms(1,1,1,13,"ORCHESTRATION-DIGITAL  |  MEDIA PLAN",bold=True,bg=navy,fg=white,align="center",size=14)

    info=brief_summary or {}
    meta=[("Date",datetime.today().strftime("%d %b %Y")),("Client",client_name),("Brand",brand_name),
          ("Campaign Name",info.get("campaign_name","")),("Campaign Start",info.get("start_date","")),
          ("Campaign End",info.get("end_date","")),("Plan Version",plan_version)]
    for i,(label,val) in enumerate(meta,2):
        ws.row_dimensions[i].height=20
        cs(i,1,label,bold=True,bg=light,fg=navy,align="right")
        ms(i,2,i,6,value=str(val))
        for c in range(7,14): cs(i,c,bg=white)

    hdr_row=10; ws.row_dimensions[hdr_row].height=40
    hdrs=["Channel","Objective","Target Audience","KPI Type","Buying Rate (LKR)","Target KPI","Spendable (USD)","Spendable (LKR)","Billable (LKR)","Creative Assets","Start Date","End Date","Days"]
    for ci,h in enumerate(hdrs,1): cs(hdr_row,ci,h,bold=True,bg=navy,fg=white,align="center",size=9)

    channel_colours={"Facebook":"1877F2","Instagram":"E1306C","YouTube":"FF0000",
                     "Google Search":"1A8A6E","Google Display":"34A853","TikTok":"010101",
                     "LinkedIn":"0A66C2","Programmatic":"7B5EA7"}
    usd_rate=320; commission=0.10; ssc_rate=0.025641; vat_rate=0.18; wht_rate=0.163
    start_str=info.get("start_date",""); end_str=info.get("end_date","")
    try:
        s=datetime.strptime(start_str,"%Y-%m-%d"); e=datetime.strptime(end_str,"%Y-%m-%d"); days=(e-s).days
    except: days=0

    current_row=hdr_row+1; channel_totals={}

    # creative_placements: dict of ch -> list of {placement, kpi_type, buying_rate, target_kpi, assets, budget}
    # If planner specified per-creative budgets, channel_kpi_data[ch] may contain a "placements" list
    for ch,sub_lkr in budget_split.items():
        colour=next((v for k,v in channel_colours.items() if k.lower() in ch.lower()),navy)
        ws.row_dimensions[current_row].height=22
        ms(current_row,1,current_row,13,ch.upper(),bold=True,bg=colour,fg=white,align="left",size=10)
        current_row+=1

        kpi=channel_kpi_data.get(ch,{})
        src_note=" ⚠️" if kpi.get("is_industry_avg") else ""
        placements=kpi.get("placements",[])  # list of per-creative placement rows

        if placements:
            # Write one row per creative placement
            for pl in placements:
                pl_lkr=pl.get("budget",round(sub_lkr/len(placements),0))
                pl_usd=round(pl_lkr/usd_rate,2)
                pl_is_meta=any(m in ch.lower() for m in ["facebook","instagram"])
                pl_billable=round(pl_lkr*1.05,0) if pl_is_meta else pl_lkr
                row_bg=light if current_row%2==0 else white
                ws.row_dimensions[current_row].height=22
                data=[pl.get("placement",ch),pl.get("objective",info.get("objective","")),
                      info.get("audience",""),pl.get("kpi_type",kpi.get("kpi_type","CPM")),
                      pl.get("buying_rate",kpi.get("buying_rate","—"))+src_note,
                      pl.get("target_kpi",kpi.get("target_kpi","—")),
                      pl_usd,pl_lkr,pl_billable,pl.get("assets",info.get("assets","")),
                      start_str,end_str,days]
                for ci,val in enumerate(data,1):
                    fmt='#,##0.00' if ci==7 else ('#,##0' if ci in (8,9) else None)
                    cs(current_row,ci,val,bg=row_bg,align="center" if ci>3 else "left",num_fmt=fmt)
                current_row+=1
            channel_totals[ch]=sub_lkr
        else:
            # Single row for channel total
            meta_channels=["facebook","instagram"]
            is_meta=any(m in ch.lower() for m in meta_channels)
            sub_usd=round(sub_lkr/usd_rate,2); billable=round(sub_lkr*1.05,0) if is_meta else sub_lkr
            row_bg=light if current_row%2==0 else white
            ws.row_dimensions[current_row].height=22
            data=[ch,info.get("objective",""),info.get("audience",""),kpi.get("kpi_type","CPM"),
                  kpi.get("buying_rate","—")+src_note,kpi.get("target_kpi","—"),
                  sub_usd,sub_lkr,billable,info.get("assets",""),start_str,end_str,days]
            for ci,val in enumerate(data,1):
                fmt='#,##0.00' if ci==7 else ('#,##0' if ci in (8,9) else None)
                cs(current_row,ci,val,bg=row_bg,align="center" if ci>3 else "left",num_fmt=fmt)
            channel_totals[ch]=sub_lkr; current_row+=1

    current_row+=1
    # 5% markup applies to Meta (Facebook/Instagram) only — all other channels billable = spendable
    meta_ch = ["facebook","instagram"]
    billable_totals = {ch: (round(v*1.05,0) if any(m in ch.lower() for m in meta_ch) else v) for ch,v in channel_totals.items()}
    total_billable = sum(billable_totals.values())
    agency_comm=round(total_billable*commission,2)
    sub1=total_billable+agency_comm
    ssc=round(sub1*ssc_rate,2)
    sub2=sub1+ssc
    vat=round(sub2*vat_rate,2)

    # WHT only on YouTube and Google billable spend
    wht_channels=["youtube","google search","google display"]
    wht_base=sum(billable_totals.get(ch,v) for ch,v in channel_totals.items() if any(w in ch.lower() for w in wht_channels))
    wht=round(wht_base*wht_rate,2) if wht_base>0 else 0
    total_invest=sub2+vat

    summary=[("Total Working Investment — Billable (LKR)",total_billable),
             ("Agency Commission (10%)",agency_comm),
             ("Sub Total",sub1),
             ("SSC Levy (2.5641%)",ssc),
             ("Sub Total",sub2),
             ("VAT (18%)",vat)]
    if wht>0:
        summary.append((f"Withholding Tax 16.3% (YouTube/Google — on LKR {wht_base:,.0f})",wht))
    summary.append(("TOTAL INVESTMENT (LKR)",total_invest))
    for label,val in summary:
        ws.row_dimensions[current_row].height=22
        is_total="TOTAL INVESTMENT" in label; is_sub=label.startswith("Sub") or label.startswith("Total W")
        bg=navy if is_total else (light if is_sub else white); fg_col=white if is_total else "1A1D23"
        ms(current_row,1,current_row,7,label,bold=is_total or is_sub,bg=bg,fg=fg_col,align="right")
        cs(current_row,8,val,bold=is_total or is_sub,bg=bg,fg=fg_col,align="right",num_fmt='#,##0.00')
        for c in range(9,14): cs(current_row,c,bg=white)
        current_row+=1

    ws3=wb.create_sheet("KPI Summary")
    for i,w in enumerate([22,18,14,22,26,20],1): ws3.column_dimensions[get_column_letter(i)].width=w
    ws3.merge_cells("A1:F1")
    h=ws3.cell(row=1,column=1,value="KPI SUMMARY BY CHANNEL")
    h.font=Font(name="Inter",bold=True,size=12,color=navy); h.fill=PatternFill("solid",fgColor=light)
    for ci,hdr in enumerate(["Channel","Budget (LKR)","KPI Type","Buying Rate","Target KPI","Data Source"],1):
        c=ws3.cell(row=2,column=ci,value=hdr); c.font=Font(name="Inter",bold=True,color=white,size=9)
        c.fill=PatternFill("solid",fgColor=navy); c.alignment=Alignment(horizontal="center",vertical="center")
    for i,(ch,budget) in enumerate(budget_split.items(),3):
        kpi=channel_kpi_data.get(ch,{})
        for ci,val in enumerate([ch,budget,kpi.get("kpi_type","—"),kpi.get("buying_rate","—"),kpi.get("target_kpi","—"),kpi.get("data_source","—")],1):
            c=ws3.cell(row=i,column=ci,value=val); c.font=Font(name="Inter",size=10)
            c.fill=PatternFill("solid",fgColor=("F0F4FA" if i%2==0 else "FFFFFF"))
            if ci==2: c.number_format='#,##0'
            c.alignment=Alignment(horizontal="center" if ci>1 else "left",vertical="center")

    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()

# ── Session defaults ──────────────────────────────────────────────────────────
for key,default in [("step",1),("brief",{}),("selected_client",None),("selected_brand",None),
                    ("combined_df",None),("generated_plan",None),("chat_messages",[]),
                    ("brief_summary",{}),("budget_split",{}),("channel_kpi_data",{}),
                    ("edit_mode",False),("edit_plan",{}),("edit_messages",[])]:
    if key not in st.session_state: st.session_state[key]=default

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='padding:16px 8px 8px 8px;'>
      <div style='font-family:Plus Jakarta Sans,sans-serif;font-size:1.3rem;font-weight:800;color:#fff;letter-spacing:-0.3px;'>🎯 Orchestration</div>
      <div style='font-size:0.72rem;color:rgba(255,255,255,0.5);margin-top:3px;letter-spacing:0.05em;text-transform:uppercase;'>Digital Planning Platform</div>
    </div>
    <hr style='margin:12px 0 16px 0;'>
    """, unsafe_allow_html=True)

    nav=st.radio("nav",[
        "📊  New Campaign Plan",
        "📁  Saved Plans",
        "📈  Data Explorer",
        "⚙️  Settings"
    ],label_visibility="collapsed")

    st.markdown("<br>",unsafe_allow_html=True)
    step=st.session_state.get("step",1)
    if nav=="📊  New Campaign Plan" and step>1:
        client_n=st.session_state.get("selected_client",{})
        brand_n=st.session_state.get("selected_brand",{})
        if client_n and brand_n:
            st.markdown(f"""
            <div style='background:rgba(255,255,255,0.1);border-radius:10px;padding:12px 14px;margin-bottom:8px;'>
              <div style='font-size:0.7rem;color:rgba(255,255,255,0.5);text-transform:uppercase;letter-spacing:0.06em;margin-bottom:4px;'>Active Session</div>
              <div style='font-size:0.85rem;font-weight:600;color:#fff;'>{client_n.get("client_name","")}</div>
              <div style='font-size:0.78rem;color:rgba(255,255,255,0.7);'>{brand_n.get("brand_name","")}</div>
              <div style='font-size:0.72rem;color:rgba(255,255,255,0.45);margin-top:4px;'>Step {step} of 4</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("""
    <div style='position:absolute;bottom:20px;left:0;right:0;padding:0 16px;'>
      <div style='font-size:0.68rem;color:rgba(255,255,255,0.3);'>POC Version 1.0 · Phase 1</div>
    </div>""", unsafe_allow_html=True)

# ── Hero ──────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero-header">
  <div class="hero-badge">Digital Planning Platform</div>
  <div class="hero-title">Orchestration-Digital</div>
  <div class="hero-sub">Plan · Analyse · Execute · Report — all in one place</div>
</div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# NEW CAMPAIGN PLAN
# ══════════════════════════════════════════════════════════════════════════════
if nav=="📊  New Campaign Plan":
    sb=get_supabase()
    steps=["1 · Client & Brand","2 · Historical Data","3 · Plan with Orchy","4 · Media Plan"]
    cols=st.columns(4)
    for i,(col,label) in enumerate(zip(cols,steps),1):
        active=st.session_state["step"]==i; done=st.session_state["step"]>i
        bg="#1e3a5f" if active else ("#1a8a6e" if done else "#e8eaf0")
        fg="#ffffff" if (active or done) else "#8a93a8"
        col.markdown(f"""<div style='background:{bg};color:{fg};border-radius:10px;padding:12px 16px;text-align:center;font-weight:600;font-size:0.82rem;'>{"✓" if done else str(i)} · {label.split("·")[1].strip()}</div>""",unsafe_allow_html=True)
    st.markdown("<br>",unsafe_allow_html=True)

    # STEP 1
    if st.session_state["step"]==1:
        st.markdown('<div class="section-header">Select Client & Brand</div>',unsafe_allow_html=True)
        clients=get_clients_list(sb); client_names=[c["client_name"] for c in clients]
        col_a,col_b=st.columns(2)
        with col_a:
            st.markdown("**Client**")
            cc=st.selectbox("Client",["— Select —","➕ Add new client"]+client_names,label_visibility="collapsed")
            if cc=="➕ Add new client":
                nc=st.text_input("New client name",placeholder="e.g. Unilever Lanka")
                if st.button("Create Client") and nc:
                    r=create_client_record(sb,nc)
                    if r: st.success(f"✅ '{nc}' created!"); st.rerun()
            elif cc!="— Select —":
                st.session_state["selected_client"]=next((c for c in clients if c["client_name"]==cc),None)
        with col_b:
            if st.session_state.get("selected_client"):
                st.markdown("**Brand**")
                brands=get_brands_for_client(sb,st.session_state["selected_client"]["id"])
                brand_names=[b["brand_name"] for b in brands]
                bc=st.selectbox("Brand",["— Select —","➕ Add new brand"]+brand_names,label_visibility="collapsed")
                if bc=="➕ Add new brand":
                    nb=st.text_input("New brand name",placeholder="e.g. Sunlight")
                    if st.button("Create Brand") and nb:
                        r=create_brand_record(sb,st.session_state["selected_client"]["id"],nb)
                        if r: st.success(f"✅ '{nb}' created!"); st.rerun()
                elif bc!="— Select —":
                    st.session_state["selected_brand"]=next((b for b in brands if b["brand_name"]==bc),None)
            else: st.info("Select a client first.")
        if st.session_state.get("selected_client") and st.session_state.get("selected_brand"):
            st.markdown("---")
            st.markdown('<div class="green-btn">',unsafe_allow_html=True)
            if st.button("Continue to Historical Data →",use_container_width=True):
                st.session_state.update({"step":2,"chat_messages":[],"generated_plan":None,"budget_split":{},"channel_kpi_data":{}})
                st.rerun()
            st.markdown('</div>',unsafe_allow_html=True)

    # STEP 2
    elif st.session_state["step"]==2:
        brand_id=st.session_state["selected_brand"]["id"]
        brand_name=st.session_state["selected_brand"]["brand_name"]
        client_name=st.session_state["selected_client"]["client_name"]
        st.markdown(f'<div class="section-header">Historical Data — {client_name} · {brand_name}</div>',unsafe_allow_html=True)

        st.markdown("""
        <div style='background:#f0f4fa;border-left:4px solid #2d5a9b;border-radius:0 10px 10px 0;padding:16px 20px;margin-bottom:20px;'>
          <div style='font-family:Plus Jakarta Sans,sans-serif;font-weight:700;color:#1e3a5f;font-size:0.92rem;margin-bottom:6px;'>📊 Why do we need your past campaign data?</div>
          <div style='font-size:0.85rem;color:#4a5168;line-height:1.7;'>
            Orchy uses your historical campaign data to make the plan <b>specific to your brand</b>, not generic.<br><br>
            From your past data, Orchy will:<br>
            &nbsp;&nbsp;• Calculate your <b>actual buying rates</b> — e.g. your real CPM on Facebook, not an industry estimate<br>
            &nbsp;&nbsp;• Set <b>realistic KPI targets</b> based on what your campaigns have actually achieved<br>
            &nbsp;&nbsp;• Identify your <b>most cost-efficient channels</b> to inform budget allocation<br><br>
            <b>What to upload:</b> Export your campaign performance reports from Meta Ads Manager, Google Ads, TikTok Ads Manager, or any other platform. CSV or Excel files work. The more historical data you provide, the more accurate the plan.
          </div>
        </div>
        """, unsafe_allow_html=True)

        saved_files=get_brand_data_files(sb,brand_id)
        if saved_files:
            total_rows=sum(f.get("row_count",0) for f in saved_files)
            st.success(f"✅ {len(saved_files)} saved file(s) for **{brand_name}** — auto-loaded.")
            c1,c2,c3=st.columns(3)
            with c1: st.markdown(f'<div class="metric-card"><div class="metric-value">{len(saved_files)}</div><div class="metric-label">Saved Files</div></div>',unsafe_allow_html=True)
            with c2: st.markdown(f'<div class="metric-card"><div class="metric-value">{total_rows:,}</div><div class="metric-label">Total Rows</div></div>',unsafe_allow_html=True)
            with c3: st.markdown(f'<div class="metric-card"><div class="metric-value">{brand_name}</div><div class="metric-label">Brand</div></div>',unsafe_allow_html=True)
            dfs=[]
            with st.expander("📁 Manage saved files"):
                for f in saved_files:
                    ca,cb,cc=st.columns([3,2,1])
                    ca.markdown(f"📄 **{f['file_name']}**")
                    cb.markdown(f"<span style='color:#8a93a8;font-size:0.82rem;'>{f.get('row_count',0):,} rows · {f['uploaded_at'][:10]}</span>",unsafe_allow_html=True)
                    if cc.button("🗑",key=f"del_{f['id']}"): delete_brand_data_file(sb,f["id"]); st.rerun()
                    try: dfs.append(pd.read_json(io.StringIO(f["data_json"])))
                    except: pass
            if dfs: st.session_state["combined_df"]=pd.concat(dfs,ignore_index=True)
        else: st.info(f"No historical data saved for **{brand_name}** yet.")
        st.markdown('<div class="section-header">Upload New Files</div>',unsafe_allow_html=True)
        uploaded_files=st.file_uploader("Drop files here",type=["csv","xlsx","xls"],accept_multiple_files=True)
        if uploaded_files:
            new_dfs=[(f.name,parse_uploaded_file(f)) for f in uploaded_files]
            new_dfs=[(n,d) for n,d in new_dfs if d is not None]
            if new_dfs:
                if st.button(f"💾 Save {len(new_dfs)} file(s) to {brand_name}",use_container_width=True):
                    for fname,df in new_dfs: save_brand_data(sb,brand_id,fname,df.to_json(),len(df))
                    st.success("✅ Saved!"); st.rerun()
                new_combined=pd.concat([d for _,d in new_dfs],ignore_index=True)
                existing=st.session_state.get("combined_df")
                st.session_state["combined_df"]=pd.concat([existing,new_combined],ignore_index=True) if existing is not None else new_combined
        st.markdown("---")
        c1,c2=st.columns(2)
        with c1:
            if st.button("← Back",use_container_width=True): st.session_state["step"]=1; st.rerun()
        with c2:
            st.markdown('<div class="green-btn">',unsafe_allow_html=True)
            if st.button("Continue to Planning Chat →",use_container_width=True):
                if st.session_state.get("combined_df") is None: st.error("Please upload at least one data file.")
                else: st.session_state["step"]=3; st.rerun()
            st.markdown('</div>',unsafe_allow_html=True)

    # STEP 3 — CHAT
    elif st.session_state["step"]==3:
        client_name=st.session_state["selected_client"]["client_name"]
        brand_name=st.session_state["selected_brand"]["brand_name"]
        data_summary=summarise_dataframe(st.session_state["combined_df"])

        st.markdown(f'<div class="section-header">Plan with Orchy — {client_name} · {brand_name}</div>',unsafe_allow_html=True)

        # Initialise with Orchy greeting
        if not st.session_state["chat_messages"]:
            with st.spinner("Orchy is getting ready…"):
                first=get_agent_response(
                    [{"role":"user","content":"Hello, I need to plan a new campaign."}],
                    client_name,brand_name,data_summary
                )
            st.session_state["chat_messages"]=[
                {"role":"user","content":"Hello, I need to plan a new campaign."},
                {"role":"agent","content":first}
            ]
            st.rerun()

        # Render all messages using native st.chat_message
        for msg in st.session_state["chat_messages"]:
            role_display = "assistant" if msg["role"]=="agent" else "user"
            avatar = "🤖" if msg["role"]=="agent" else "🧑‍💼"
            with st.chat_message(role_display, avatar=avatar):
                st.markdown(msg["content"].replace("[BRIEF_COMPLETE]","").strip())

        last_agent=next((m["content"] for m in reversed(st.session_state["chat_messages"]) if m["role"]=="agent"),"")
        brief_auto_complete="[BRIEF_COMPLETE]" in last_agent
        num_user_msgs=len([m for m in st.session_state["chat_messages"] if m["role"]=="user"])

        # Generate button above input when ready
        if num_user_msgs>=3:
            if brief_auto_complete:
                st.success("✅ Orchy has all the information needed. Ready to generate your media plan!")
            else:
                st.info("💡 Ready to generate? Click below — or keep chatting to add more detail.")
            st.markdown('<div class="green-btn">',unsafe_allow_html=True)
            if st.button("🚀 Generate Media Plan",use_container_width=True):
                with st.spinner("🤖 Calculating budget split, buying rates and KPI targets…"):
                    try:
                        brief_summary=extract_brief_from_conversation(st.session_state["chat_messages"])
                        st.session_state["brief_summary"]=brief_summary
                        benchmarks=extract_channel_benchmarks(st.session_state["combined_df"])
                        channels=brief_summary.get("channels",[])
                        audience_sizes=brief_summary.get("audience_sizes",{})
                        total_budget=float(brief_summary.get("total_budget",0))
                        objective=brief_summary.get("objective","Brand Awareness")
                        budget_split=calculate_budget_split(channels,total_budget,objective,audience_sizes,benchmarks)
                        st.session_state["budget_split"]=budget_split
                        channel_placements=brief_summary.get("channel_placements",{})
                        channel_kpi_data={}
                        for ch,b in budget_split.items():
                            kpi=calculate_channel_kpis(ch,b,benchmarks,objective)
                            # Attach per-creative placements if specified
                            if ch in channel_placements and channel_placements[ch]:
                                enriched_placements=[]
                                for pl in channel_placements[ch]:
                                    pl_budget=pl.get("budget",0) or round(b/len(channel_placements[ch]),0)
                                    pl_kpi=calculate_channel_kpis(ch,pl_budget,benchmarks,pl.get("objective",objective))
                                    enriched_placements.append({
                                        "placement": pl.get("placement",""),
                                        "kpi_type":  pl.get("kpi_type",pl_kpi["kpi_type"]),
                                        "buying_rate": pl_kpi["buying_rate"],
                                        "target_kpi":  pl_kpi["target_kpi"],
                                        "assets":    pl.get("assets",""),
                                        "objective": pl.get("objective",objective),
                                        "budget":    pl_budget,
                                    })
                                kpi["placements"]=enriched_placements
                            channel_kpi_data[ch]=kpi
                        st.session_state["channel_kpi_data"]=channel_kpi_data
                        data_gaps=get_data_gaps(channels,benchmarks)
                        plan_text=generate_media_plan(st.session_state["chat_messages"],client_name,brand_name,data_summary,budget_split,channel_kpi_data,data_gaps)
                        st.session_state["generated_plan"]=plan_text
                        st.session_state["step"]=4
                        st.rerun()
                    except Exception as e: st.error(f"Error: {e}")
            st.markdown('</div>',unsafe_allow_html=True)
            st.markdown("---")

        # Native chat input — Enter to send, Shift+Enter for new line, clears automatically
        if user_input := st.chat_input("Message Orchy… (Enter to send, Shift+Enter for new line)"):
            st.session_state["chat_messages"].append({"role":"user","content":user_input.strip()})
            with st.chat_message("user", avatar="🧑‍💼"):
                st.markdown(user_input.strip())
            with st.chat_message("assistant", avatar="🤖"):
                with st.spinner("Orchy is thinking…"):
                    reply=get_agent_response(st.session_state["chat_messages"],client_name,brand_name,data_summary)
                st.markdown(reply.replace("[BRIEF_COMPLETE]","").strip())
            st.session_state["chat_messages"].append({"role":"agent","content":reply})
            st.rerun()

        if st.button("← Back to Data"): st.session_state["step"]=2; st.rerun()

    # STEP 4
    elif st.session_state["step"]==4:
        client_name=st.session_state["selected_client"]["client_name"]
        brand_name=st.session_state["selected_brand"]["brand_name"]
        plan_text=st.session_state["generated_plan"]
        brief_summary=st.session_state.get("brief_summary",{})
        budget_split=st.session_state.get("budget_split",{})
        channel_kpi_data=st.session_state.get("channel_kpi_data",{})
        total_budget=float(brief_summary.get("total_budget",0))

        st.markdown(f'<div class="section-header">{client_name} · {brand_name} · {brief_summary.get("campaign_name","Media Plan")}</div>',unsafe_allow_html=True)

        try:
            s=datetime.strptime(brief_summary.get("start_date",""),"%Y-%m-%d")
            e=datetime.strptime(brief_summary.get("end_date",""),"%Y-%m-%d"); days=(e-s).days
        except: days="—"

        c1,c2,c3,c4=st.columns(4)
        with c1: st.markdown(f'<div class="metric-card"><div class="metric-value">{format_lkr(total_budget)}</div><div class="metric-label">Total Budget</div></div>',unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-card"><div class="metric-value">{days} days</div><div class="metric-label">Duration</div></div>',unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-card"><div class="metric-value">{len(budget_split)}</div><div class="metric-label">Channels</div></div>',unsafe_allow_html=True)
        with c4:
            obj=brief_summary.get("objective","—")
            st.markdown(f'<div class="metric-card"><div class="metric-value">{obj.split()[0] if obj else "—"}</div><div class="metric-label">Objective</div></div>',unsafe_allow_html=True)

        # Data gaps warning
        benchmarks=extract_channel_benchmarks(st.session_state["combined_df"]) if st.session_state.get("combined_df") is not None else {}
        gaps=get_data_gaps(list(budget_split.keys()),benchmarks)
        if gaps:
            st.warning("⚠️ **Industry averages used for the following channels** (no historical data found):\n\n"+"\n\n".join(gaps))

        if budget_split:
            st.markdown('<div class="section-header">Budget Split & KPI Targets</div>',unsafe_allow_html=True)
            rows=[]
            for ch,budget in budget_split.items():
                kpi=channel_kpi_data.get(ch,{})
                rows.append({"Channel":ch,"Budget (LKR)":f"LKR {budget:,.0f}",
                             "% of Total":f"{budget/total_budget*100:.1f}%" if total_budget>0 else "—",
                             "KPI Type":kpi.get("kpi_type","—"),"Buying Rate":kpi.get("buying_rate","—"),
                             "Target KPI":kpi.get("target_kpi","—"),"Data Source":kpi.get("data_source","—")})
            st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)

        # ── Action buttons at the TOP ────────────────────────────────────────
        pv=get_plan_version(sb,brief_summary.get("campaign_name",""),st.session_state["selected_brand"]["id"])
        excel_bytes=build_excel(brief_summary,plan_text,client_name,brand_name,budget_split,channel_kpi_data,pv)
        ca,cb,cc,cd=st.columns(4)
        with ca:
            st.download_button("📥 Download TXT",data=plan_text,
                file_name=f"{brief_summary.get('campaign_name','plan').replace(' ','_')}_plan.txt",
                mime="text/plain",use_container_width=True)
        with cb:
            st.download_button("📊 Download Excel",data=excel_bytes,
                file_name=f"{brief_summary.get('campaign_name','plan').replace(' ','_')}_MediaPlan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        with cc:
            if st.button("💾 Save to Library",use_container_width=True):
                save_plan(sb,{"brand_id":st.session_state["selected_brand"]["id"],"client_name":client_name,
                    "brand_name":brand_name,"campaign_name":brief_summary.get("campaign_name",""),
                    "objective":brief_summary.get("objective",""),"total_budget":total_budget,
                    "start_date":brief_summary.get("start_date",""),"end_date":brief_summary.get("end_date",""),
                    "channels":json.dumps(list(budget_split.keys())),"market":brief_summary.get("market","Sri Lanka"),
                    "kpi_focus":"","plan_text":plan_text,"plan_version":pv,
                    "budget_split":json.dumps(budget_split),"channel_kpi_data":json.dumps(channel_kpi_data),
                    "created_at":datetime.utcnow().isoformat()})
                st.success(f"Saved as {pv}!")
        with cd:
            st.markdown('<div class="green-btn">',unsafe_allow_html=True)
            if st.button("➕ New Plan",use_container_width=True):
                # Full reset — clear everything including client/brand selection
                for k in ["step","brief","selected_client","selected_brand","combined_df",
                          "generated_plan","chat_messages","brief_summary",
                          "budget_split","channel_kpi_data","edit_mode","edit_plan","edit_messages"]:
                    st.session_state[k] = 1 if k=="step" else ([] if k in ["chat_messages","edit_messages"] else ({} if k not in ["selected_client","selected_brand","combined_df","generated_plan"] else None))
                st.rerun()
            st.markdown('</div>',unsafe_allow_html=True)

        st.markdown("---")

        # ── Plan rendered as structured sections with tables ──────────────────
        st.markdown('<div class="section-header">Media Plan</div>',unsafe_allow_html=True)

        # Split plan into sections and render with st.markdown (handles tables natively)
        sections = plan_text.split("## ")
        if len(sections) <= 1:
            # Fallback if no ## headers — render in expander to keep it compact
            with st.expander("📄 View Full Plan", expanded=True):
                st.markdown(plan_text)
        else:
            for section in sections:
                if not section.strip(): continue
                lines = section.strip().split("\n",1)
                title = lines[0].strip()
                body = lines[1].strip() if len(lines)>1 else ""
                # Show exec summary expanded, rest collapsed
                expanded = any(k in title.lower() for k in ["executive","summary","channel plan","kpi"])
                with st.expander(f"**{title}**", expanded=expanded):
                    st.markdown(body)

# ══════════════════════════════════════════════════════════════════════════════
# SAVED PLANS
# ══════════════════════════════════════════════════════════════════════════════
elif nav=="📁  Saved Plans":
    sb=get_supabase()

    # Edit mode
    if st.session_state.get("edit_mode") and st.session_state.get("edit_plan"):
        ep=st.session_state["edit_plan"]
        client_name=ep.get("client_name","")
        brand_name=ep.get("brand_name","")
        original_plan=ep.get("plan_text","")
        budget_split=json.loads(ep.get("budget_split","{}"))
        channel_kpi_data=json.loads(ep.get("channel_kpi_data","{}"))

        st.markdown(f'<div class="section-header">✏️ Editing — {client_name} · {brand_name} · {ep.get("campaign_name","")} {ep.get("plan_version","")}</div>',unsafe_allow_html=True)
        st.caption("Tell Orchy what you'd like to change. Be specific about budgets, channels, or strategy adjustments.")

        if not st.session_state["edit_messages"]:
            with st.spinner("Orchy is reviewing the plan…"):
                first=get_agent_response(
                    [{"role":"user","content":"I need to edit this campaign plan. Please review it and ask me what I'd like to change."}],
                    client_name,brand_name,"",mode="editing",existing_plan=original_plan
                )
            st.session_state["edit_messages"]=[
                {"role":"user","content":"I need to edit this campaign plan. Please review it and ask me what I'd like to change."},
                {"role":"agent","content":first}
            ]
            st.rerun()

        # Render edit chat using native components
        for msg in st.session_state["edit_messages"]:
            role_display="assistant" if msg["role"]=="agent" else "user"
            avatar="🤖" if msg["role"]=="agent" else "🧑‍💼"
            with st.chat_message(role_display,avatar=avatar):
                st.markdown(msg["content"].replace("[EDIT_COMPLETE]","").strip())

        last_agent=next((m["content"] for m in reversed(st.session_state["edit_messages"]) if m["role"]=="agent"),"")
        edit_complete="[EDIT_COMPLETE]" in last_agent
        num_edit_msgs=len([m for m in st.session_state["edit_messages"] if m["role"]=="user"])

        # Native chat input for edits
        if edit_input := st.chat_input("Tell Orchy what to change… (Enter to send)"):
            st.session_state["edit_messages"].append({"role":"user","content":edit_input.strip()})
            with st.chat_message("user",avatar="🧑‍💼"):
                st.markdown(edit_input.strip())
            with st.chat_message("assistant",avatar="🤖"):
                with st.spinner("Orchy is thinking…"):
                    reply=get_agent_response(st.session_state["edit_messages"],client_name,brand_name,"",mode="editing",existing_plan=original_plan)
                st.markdown(reply.replace("[EDIT_COMPLETE]","").strip())
            st.session_state["edit_messages"].append({"role":"agent","content":reply})
            st.rerun()

        if num_edit_msgs>=1:
            st.markdown("---")
            if edit_complete:
                st.success("✅ Orchy has noted all the changes. Ready to generate the updated plan!")
            else:
                st.info("💡 Keep adding instructions, or click below when ready to regenerate.")
            st.markdown('<div class="green-btn">',unsafe_allow_html=True)
            if st.button("🚀 Apply Edits & Generate New Version",use_container_width=True):
                with st.spinner("🤖 Applying edits and generating updated plan…"):
                    try:
                        new_plan=apply_plan_edits(st.session_state["edit_messages"],original_plan,client_name,brand_name,budget_split,channel_kpi_data)
                        new_version=get_plan_version(sb,ep.get("campaign_name",""),ep.get("brand_id",""))
                        save_plan(sb,{"brand_id":ep.get("brand_id"),"client_name":client_name,"brand_name":brand_name,
                            "campaign_name":ep.get("campaign_name",""),"objective":ep.get("objective",""),
                            "total_budget":ep.get("total_budget",0),"start_date":ep.get("start_date",""),
                            "end_date":ep.get("end_date",""),"channels":ep.get("channels","[]"),
                            "market":ep.get("market","Sri Lanka"),"kpi_focus":"","plan_text":new_plan,
                            "plan_version":new_version,"budget_split":ep.get("budget_split","{}"),
                            "channel_kpi_data":ep.get("channel_kpi_data","{}"),
                            "created_at":datetime.utcnow().isoformat()})
                        st.success(f"✅ New version {new_version} saved to library!")
                        st.markdown('<div class="plan-card">'+new_plan.replace("\n","<br>")+"</div>",unsafe_allow_html=True)
                        st.session_state.update({"edit_mode":False,"edit_plan":{},"edit_messages":[]})
                    except Exception as e: st.error(f"Error: {e}")
            st.markdown('</div>',unsafe_allow_html=True)

        st.markdown("---")
        if st.button("← Cancel Edit"):
            st.session_state.update({"edit_mode":False,"edit_plan":{},"edit_messages":[]})
            st.rerun()

    else:
        st.markdown('<div class="section-header">Saved Campaign Plans</div>',unsafe_allow_html=True)
        try:
            plans=load_saved_plans(sb)
            if not plans:
                st.info("No saved plans yet. Create your first plan to see it here.")
            else:
                # Group by campaign
                for plan in plans:
                    budget=float(plan.get("total_budget",0) or 0)
                    version=plan.get("plan_version","V1.0")
                    label=f"📋 {plan.get('client_name','—')} · {plan.get('brand_name','—')} · {plan.get('campaign_name','Unnamed')} [{version}] — {format_lkr(budget)}"
                    with st.expander(label):
                        c1,c2,c3,c4=st.columns(4)
                        c1.metric("Objective",plan.get("objective","—"))
                        c2.metric("Budget",format_lkr(budget))
                        c3.metric("Version",version)
                        c4.metric("Market",plan.get("market","—"))
                        st.markdown(f"**Dates:** {plan.get('start_date','')} → {plan.get('end_date','')}")

                        # KPI table if available
                        if plan.get("channel_kpi_data"):
                            try:
                                kpi_data=json.loads(plan["channel_kpi_data"])
                                budget_data=json.loads(plan.get("budget_split","{}"))
                                if kpi_data:
                                    rows=[{"Channel":ch,"Budget":format_lkr(budget_data.get(ch,0)),
                                           "KPI Type":v.get("kpi_type","—"),"Buying Rate":v.get("buying_rate","—"),
                                           "Target KPI":v.get("target_kpi","—")} for ch,v in kpi_data.items()]
                                    st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)
                            except: pass

                        st.markdown("---")
                        st.markdown(plan.get("plan_text",""))
                        st.markdown("---")

                        col_edit,col_dl=st.columns(2)
                        with col_edit:
                            if st.button(f"✏️ Edit this Plan",key=f"edit_{plan['id']}",use_container_width=True):
                                st.session_state.update({"edit_mode":True,"edit_plan":plan,"edit_messages":[]})
                                st.rerun()
                        with col_dl:
                            try:
                                bs=json.loads(plan.get("budget_split","{}"))
                                ck=json.loads(plan.get("channel_kpi_data","{}"))
                                bs_obj={"campaign_name":plan.get("campaign_name",""),"objective":plan.get("objective",""),
                                        "start_date":plan.get("start_date",""),"end_date":plan.get("end_date",""),"audience":""}
                                excel=build_excel(bs_obj,plan.get("plan_text",""),plan.get("client_name",""),plan.get("brand_name",""),bs,ck,version)
                                st.download_button("📊 Download Excel",data=excel,
                                    file_name=f"{plan.get('campaign_name','plan').replace(' ','_')}_{version}_MediaPlan.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True,key=f"dl_{plan['id']}")
                            except: pass
        except Exception as e: st.error(f"Could not load plans: {e}")

# ══════════════════════════════════════════════════════════════════════════════
# DATA EXPLORER
# ══════════════════════════════════════════════════════════════════════════════
elif nav=="📈  Data Explorer":
    st.markdown('<div class="section-header">Data Explorer</div>',unsafe_allow_html=True)
    if st.session_state.get("combined_df") is None:
        st.info("Complete Step 2 in New Campaign Plan to load data here.")
    else:
        df=st.session_state["combined_df"]
        st.caption(f"{len(df):,} rows · {len(df.columns)} columns")
        st.dataframe(df,use_container_width=True)
        numeric_cols=df.select_dtypes(include='number').columns.tolist()
        if len(numeric_cols)>=2:
            st.markdown('<div class="section-header">Quick Chart</div>',unsafe_allow_html=True)
            c1,c2=st.columns(2)
            x_col=c1.selectbox("X axis",df.columns.tolist())
            y_col=c2.selectbox("Y axis",numeric_cols)
            chart_type=st.radio("Chart type",["Line","Bar","Area"],horizontal=True)
            chart_df=df[[x_col,y_col]].dropna().set_index(x_col)
            if chart_type=="Line": st.line_chart(chart_df)
            elif chart_type=="Bar": st.bar_chart(chart_df)
            else: st.area_chart(chart_df)

# ══════════════════════════════════════════════════════════════════════════════
# SETTINGS
# ══════════════════════════════════════════════════════════════════════════════
elif nav=="⚙️  Settings":
    sb=get_supabase()
    st.markdown('<div class="section-header">Platform Settings</div>',unsafe_allow_html=True)
    st.caption("Configure platform efficiency benchmarks used to optimise budget allocation. Audience sizes are captured per campaign during the Orchy planning chat.")

    platforms=[
        {"name":"Facebook","default_freq_awareness":3,"default_freq_conversion":7,"reach_curve_inflection":0.4},
        {"name":"Instagram","default_freq_awareness":3,"default_freq_conversion":6,"reach_curve_inflection":0.35},
        {"name":"YouTube","default_freq_awareness":4,"default_freq_conversion":5,"reach_curve_inflection":0.45},
        {"name":"Google Search","default_freq_awareness":1,"default_freq_conversion":3,"reach_curve_inflection":0.6},
        {"name":"Google Display","default_freq_awareness":5,"default_freq_conversion":8,"reach_curve_inflection":0.3},
        {"name":"TikTok","default_freq_awareness":3,"default_freq_conversion":5,"reach_curve_inflection":0.38},
        {"name":"LinkedIn","default_freq_awareness":4,"default_freq_conversion":6,"reach_curve_inflection":0.5},
        {"name":"Programmatic","default_freq_awareness":6,"default_freq_conversion":9,"reach_curve_inflection":0.3},
    ]
    existing_settings=get_platform_settings(sb)
    existing_dict={p["platform_name"]:p for p in existing_settings}

    st.markdown('<div class="section-header">Frequency Caps & Reach Curve Parameters</div>',unsafe_allow_html=True)
    st.caption("Frequency cap = max times a user sees your ad per campaign. Reach curve inflection = point (as % of targetable audience) where diminishing returns begin.")

    for p in platforms:
        pname=p["name"]; saved=existing_dict.get(pname,{})
        with st.expander(f"⚙️ {pname}"):
            col1,col2,col3=st.columns(3)
            fa=col1.number_input("Freq Cap — Awareness",min_value=1,max_value=20,value=int(saved.get("freq_cap_awareness",p["default_freq_awareness"])),key=f"fa_{pname}")
            fc=col2.number_input("Freq Cap — Conversion",min_value=1,max_value=20,value=int(saved.get("freq_cap_conversion",p["default_freq_conversion"])),key=f"fc_{pname}")
            ri=col3.number_input("Reach Curve Inflection (%)",min_value=0.1,max_value=1.0,step=0.05,value=float(saved.get("reach_curve_inflection",p["reach_curve_inflection"])),key=f"ri_{pname}")
            notes=st.text_input("Notes",value=saved.get("notes",""),key=f"n_{pname}",placeholder="e.g. Strong for video in LK, peak hours 7–10pm")
            if st.button(f"Save {pname}",key=f"save_{pname}"):
                save_platform_setting(sb,pname,{"freq_cap_awareness":fa,"freq_cap_conversion":fc,"reach_curve_inflection":ri,"notes":notes})
                st.success(f"✅ {pname} settings saved!")

    st.markdown('<div class="section-header">Sri Lanka Industry Averages (Fallback)</div>',unsafe_allow_html=True)
    st.caption("These are used when no historical data exists for a channel. Clearly marked in plan outputs.")
    rows=[{"Platform":k,"CPM (LKR)":v.get("cpm","—"),"CPC (LKR)":v.get("cpc","—"),"CPV (LKR)":v.get("cpv","—"),"CTR (%)":v.get("ctr","—"),"ROAS":v.get("roas","—")} for k,v in INDUSTRY_AVERAGES.items()]
    st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)

    st.markdown('<div class="section-header">About Budget Optimisation</div>',unsafe_allow_html=True)
    st.markdown("""<div class="settings-card"><p style='font-size:0.88rem;color:#4a5168;line-height:1.7;'>
    Budget is allocated using a multi-factor model:<br><br>
    <b>1. Objective weighting</b> — channels scored by how well they match the campaign objective.<br>
    <b>2. Targetable audience size</b> — captured during the Orchy chat from each platform's estimator. Larger audiences allow more budget before diminishing returns.<br>
    <b>3. Historical cost efficiency</b> — lower historical CPM/CPC = more efficient = higher budget score.<br>
    <b>4. Reach curve</b> — the inflection point controls when diminishing returns kick in. Minimum 5% budget floor per channel.<br>
    <b>5. Data transparency</b> — if no historical data exists for a channel, Sri Lanka industry averages are used and clearly flagged.<br><br>
    <i>In Phase 2, audience sizes and buying rates will be pulled directly from platform APIs.</i>
    </p></div>""", unsafe_allow_html=True)

# Run SQL for new column needed
# ALTER TABLE campaign_plans ADD COLUMN IF NOT EXISTS plan_version text DEFAULT 'V1.0';
# ALTER TABLE campaign_plans ADD COLUMN IF NOT EXISTS budget_split text;
# ALTER TABLE campaign_plans ADD COLUMN IF NOT EXISTS channel_kpi_data text;