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
  html, body, [class*="css"] { font-family:'Inter',sans-serif; background-color:#ffffff !important; color:#1a1d23; }
  .stApp { background:#f8f9fc !important; }
  [data-testid="stSidebar"] { background:#ffffff !important; border-right:1px solid #e8eaf0 !important; }
  [data-testid="stSidebar"] * { color:#1a1d23 !important; }
  .hero-header { background:linear-gradient(135deg,#1e3a5f 0%,#2d5a9b 50%,#1a8a6e 100%); border-radius:16px; padding:36px 40px; margin-bottom:28px; }
  .hero-badge { display:inline-block; background:rgba(255,255,255,0.15); color:#fff; border:1px solid rgba(255,255,255,0.25); border-radius:20px; font-size:0.7rem; font-weight:600; letter-spacing:0.1em; text-transform:uppercase; padding:4px 14px; margin-bottom:12px; }
  .hero-title { font-family:'Plus Jakarta Sans',sans-serif; font-size:2rem; font-weight:800; color:#fff; letter-spacing:-0.5px; margin:0 0 8px 0; }
  .hero-sub { font-size:0.9rem; color:rgba(255,255,255,0.75); margin:0; }
  .metric-card { background:#fff; border:1px solid #e8eaf0; border-radius:12px; padding:20px 24px; text-align:center; box-shadow:0 1px 4px rgba(0,0,0,0.05); }
  .metric-value { font-family:'Plus Jakarta Sans',sans-serif; font-size:1.6rem; font-weight:700; color:#1e3a5f; }
  .metric-label { font-size:0.75rem; color:#8a93a8; text-transform:uppercase; letter-spacing:0.06em; margin-top:4px; }
  .section-header { font-family:'Plus Jakarta Sans',sans-serif; font-size:1rem; font-weight:700; color:#1a1d23; border-left:3px solid #2d5a9b; padding-left:12px; margin:24px 0 14px 0; }
  .chat-container { background:#fff; border:1px solid #e8eaf0; border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.05); margin-bottom:16px; }
  .chat-header { background:linear-gradient(135deg,#1e3a5f,#2d5a9b); padding:16px 20px; }
  .chat-header-title { color:#fff; font-family:'Plus Jakarta Sans',sans-serif; font-weight:700; font-size:0.95rem; margin:0; }
  .chat-header-sub { color:rgba(255,255,255,0.7); font-size:0.78rem; margin:2px 0 0 0; }
  .chat-messages { padding:20px; max-height:500px; overflow-y:auto; }
  .msg-agent { display:flex; gap:12px; margin-bottom:16px; align-items:flex-start; }
  .msg-user { display:flex; gap:12px; margin-bottom:16px; align-items:flex-start; flex-direction:row-reverse; }
  .avatar-agent { width:32px; height:32px; background:linear-gradient(135deg,#1e3a5f,#2d5a9b); border-radius:50%; display:flex; align-items:center; justify-content:center; color:#fff; font-size:0.75rem; font-weight:700; flex-shrink:0; }
  .avatar-user { width:32px; height:32px; background:linear-gradient(135deg,#1a8a6e,#22b894); border-radius:50%; display:flex; align-items:center; justify-content:center; color:#fff; font-size:0.75rem; font-weight:700; flex-shrink:0; }
  .bubble-agent { background:#f0f4fa; border-radius:0 12px 12px 12px; padding:12px 16px; max-width:80%; font-size:0.88rem; line-height:1.6; color:#1a1d23; }
  .bubble-user { background:linear-gradient(135deg,#1e3a5f,#2d5a9b); border-radius:12px 0 12px 12px; padding:12px 16px; max-width:80%; font-size:0.88rem; line-height:1.6; color:#fff; }
  .plan-card { background:#fff; border:1px solid #e8eaf0; border-radius:14px; padding:28px 32px; margin-top:16px; line-height:1.8; color:#2d3341; box-shadow:0 1px 4px rgba(0,0,0,0.05); font-size:0.92rem; }
  .settings-card { background:#fff; border:1px solid #e8eaf0; border-radius:12px; padding:20px 24px; margin-bottom:16px; box-shadow:0 1px 4px rgba(0,0,0,0.05); }
  .stButton > button { background:linear-gradient(135deg,#1e3a5f,#2d5a9b) !important; color:white !important; border:none !important; border-radius:8px !important; font-weight:600 !important; font-size:0.9rem !important; padding:10px 24px !important; }
  .stButton > button:hover { opacity:0.88 !important; }
  .green-btn .stButton > button { background:linear-gradient(135deg,#1a8a6e,#22b894) !important; }
  .stTextInput > div > div > input, .stNumberInput > div > div > input,
  .stTextArea textarea, .stDateInput > div > div > input { background:#fff !important; border:1px solid #d0d5e0 !important; border-radius:8px !important; color:#1a1d23 !important; font-size:0.9rem !important; }
  .stSelectbox > div > div { background:#fff !important; border:1px solid #d0d5e0 !important; border-radius:8px !important; color:#1a1d23 !important; }
  .stTabs [data-baseweb="tab-list"] { background:#eef1f8; border-radius:10px; padding:4px; gap:4px; }
  .stTabs [data-baseweb="tab"] { background:transparent !important; color:#6b7590 !important; border-radius:7px !important; font-weight:500 !important; font-size:0.88rem !important; }
  .stTabs [aria-selected="true"] { background:#fff !important; color:#1e3a5f !important; font-weight:600 !important; box-shadow:0 1px 4px rgba(0,0,0,0.08) !important; }
  div[data-testid="stExpander"] { background:#fff !important; border:1px solid #e8eaf0 !important; border-radius:10px !important; }
  label { color:#4a5168 !important; font-size:0.85rem !important; font-weight:500 !important; }
  #MainMenu, footer, header { visibility:hidden; }
</style>
""", unsafe_allow_html=True)

# ── Supabase ──────────────────────────────────────────────────────────────────
@st.cache_resource
def get_supabase() -> Client:
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
    except Exception as e:
        st.error(f"Error creating client: {e}"); return None

def create_brand_record(sb, client_id, name):
    try:
        r = sb.table("brands").insert({"client_id": client_id, "brand_name": name, "created_at": datetime.utcnow().isoformat()}).execute()
        return r.data[0] if r.data else None
    except Exception as e:
        st.error(f"Error creating brand: {e}"); return None

def get_brand_data_files(sb, brand_id):
    try: return sb.table("brand_data").select("*").eq("brand_id", brand_id).order("uploaded_at", desc=True).execute().data or []
    except: return []

def save_brand_data(sb, brand_id, file_name, data_json, row_count):
    try:
        sb.table("brand_data").insert({"brand_id": brand_id, "file_name": file_name, "data_json": data_json, "row_count": row_count, "uploaded_at": datetime.utcnow().isoformat()}).execute()
        return True
    except Exception as e:
        st.error(f"Error saving data: {e}"); return False

def delete_brand_data_file(sb, file_id):
    try: sb.table("brand_data").delete().eq("id", file_id).execute(); return True
    except: return False

def save_plan(sb, plan_data):
    try: sb.table("campaign_plans").insert(plan_data).execute(); return True
    except Exception as e: st.warning(f"Could not save: {e}"); return False

def load_saved_plans(sb):
    try: return sb.table("campaign_plans").select("*").order("created_at", desc=True).limit(30).execute().data or []
    except: return []

def get_platform_settings(sb):
    try: return sb.table("platform_settings").select("*").order("platform_name").execute().data or []
    except: return []

def save_platform_setting(sb, platform_name, data):
    try:
        existing = sb.table("platform_settings").select("id").eq("platform_name", platform_name).execute().data
        if existing:
            sb.table("platform_settings").update({**data, "updated_at": datetime.utcnow().isoformat()}).eq("platform_name", platform_name).execute()
        else:
            sb.table("platform_settings").insert({"platform_name": platform_name, **data, "updated_at": datetime.utcnow().isoformat()}).execute()
        return True
    except Exception as e:
        st.error(f"Error saving setting: {e}"); return False

# ── Data helpers ──────────────────────────────────────────────────────────────
def parse_uploaded_file(f):
    n = f.name.lower()
    if n.endswith(".csv"): return pd.read_csv(f)
    elif n.endswith((".xlsx", ".xls")): return pd.read_excel(f)
    return None

def extract_channel_benchmarks(df: pd.DataFrame) -> dict:
    """
    Scan historical data and extract average buying rates per channel.
    Returns dict: { channel -> { kpi_type, avg_rate, ctr, cpm, cpc, cpv, cpa, roas } }
    """
    benchmarks = {}
    col_lower = {c.lower(): c for c in df.columns}

    def avg(col_key):
        matches = [c for cl, c in col_lower.items() if col_key in cl]
        if matches:
            try:
                vals = pd.to_numeric(df[matches[0]], errors="coerce").dropna()
                return round(float(vals.mean()), 2) if len(vals) > 0 else None
            except: return None
        return None

    # Try to split by channel/campaign if column exists
    channel_col = next((c for cl, c in col_lower.items() if "channel" in cl or "platform" in cl or "campaign" in cl), None)

    platform_keywords = {
        "Facebook":      ["facebook", "fb", "meta"],
        "Instagram":     ["instagram", "ig"],
        "TikTok":        ["tiktok", "tik tok"],
        "YouTube":       ["youtube", "yt"],
        "Google Search": ["google search", "search", "sem"],
        "Google Display":["google display", "display", "gdn"],
        "LinkedIn":      ["linkedin"],
        "Programmatic":  ["programmatic", "dsp", "trade desk"],
    }

    if channel_col:
        for platform, keywords in platform_keywords.items():
            mask = df[channel_col].astype(str).str.lower().apply(
                lambda x: any(kw in x for kw in keywords)
            )
            sub = df[mask]
            if len(sub) > 0:
                sub_col = {c.lower(): c for c in sub.columns}
                def sub_avg(key):
                    matches = [c for cl, c in sub_col.items() if key in cl]
                    if matches:
                        try:
                            vals = pd.to_numeric(sub[matches[0]], errors="coerce").dropna()
                            return round(float(vals.mean()), 2) if len(vals) > 0 else None
                        except: return None
                    return None
                benchmarks[platform] = {
                    "cpm":  sub_avg("cpm"),
                    "cpc":  sub_avg("cpc"),
                    "cpv":  sub_avg("cpv"),
                    "cpa":  sub_avg("cpa"),
                    "ctr":  sub_avg("ctr"),
                    "roas": sub_avg("roas"),
                    "rows": len(sub)
                }
    else:
        # No channel column — use global averages as fallback
        global_stats = {
            "cpm":  avg("cpm"),
            "cpc":  avg("cpc"),
            "cpv":  avg("cpv"),
            "cpa":  avg("cpa"),
            "ctr":  avg("ctr"),
            "roas": avg("roas"),
            "rows": len(df)
        }
        for platform in platform_keywords:
            benchmarks[platform] = global_stats.copy()

    return benchmarks

def calculate_budget_split(channels, total_budget, objective, audience_sizes, benchmarks, platform_settings):
    """
    Calculate optimised budget split based on:
    - Targetable audience size per platform (from Orchy chat)
    - Reach curve efficiency (diminishing returns model)
    - Historical cost efficiency per channel
    - Objective weighting
    Returns dict: { channel -> budget_lkr }
    """
    # Objective weights — which channels are prioritised per objective
    objective_weights = {
        "Brand Awareness":        {"Facebook":1.3,"Instagram":1.2,"YouTube":1.4,"Google Display":1.0,"TikTok":1.1,"Google Search":0.6,"LinkedIn":0.8,"Programmatic":0.9},
        "Reach & Frequency":      {"Facebook":1.4,"Instagram":1.3,"YouTube":1.3,"Google Display":0.9,"TikTok":1.0,"Google Search":0.5,"LinkedIn":0.7,"Programmatic":0.8},
        "Video Views":            {"YouTube":1.5,"TikTok":1.4,"Facebook":1.1,"Instagram":1.2,"Google Display":0.8,"Google Search":0.4,"LinkedIn":0.7,"Programmatic":0.9},
        "Website Traffic":        {"Google Search":1.5,"Facebook":1.1,"Instagram":1.0,"TikTok":0.9,"Google Display":1.0,"YouTube":0.8,"LinkedIn":1.0,"Programmatic":0.9},
        "Lead Generation":        {"Google Search":1.4,"Facebook":1.3,"LinkedIn":1.5,"Instagram":1.1,"TikTok":0.8,"YouTube":0.7,"Google Display":0.9,"Programmatic":0.8},
        "E-commerce / Conversions":{"Google Search":1.5,"Facebook":1.3,"Instagram":1.2,"TikTok":1.0,"YouTube":0.9,"Google Display":1.0,"LinkedIn":0.7,"Programmatic":1.1},
        "App Installs":           {"Facebook":1.3,"Instagram":1.2,"TikTok":1.4,"Google Search":1.1,"YouTube":1.0,"Google Display":0.9,"LinkedIn":0.7,"Programmatic":0.9},
        "Engagement":             {"Instagram":1.4,"TikTok":1.5,"Facebook":1.2,"YouTube":1.0,"LinkedIn":0.9,"Google Search":0.5,"Google Display":0.7,"Programmatic":0.7},
    }
    obj_key = next((k for k in objective_weights if k.lower() in objective.lower()), "Brand Awareness")
    weights = objective_weights[obj_key]

    scores = {}
    for ch in channels:
        # Match channel to weight key
        wt = next((v for k, v in weights.items() if k.lower() in ch.lower()), 1.0)

        # Audience size score (log scale — bigger audience = more budget headroom)
        aud = audience_sizes.get(ch, 500000)
        aud_score = np.log10(max(aud, 1000)) / np.log10(10_000_000)

        # Historical efficiency score (lower CPM/CPC = more efficient)
        bench = benchmarks.get(ch, {})
        cpm = bench.get("cpm") or 500
        efficiency_score = min(1000 / max(cpm, 50), 2.0)  # cap at 2x

        # Reach curve — diminishing returns after ~40% of audience reached
        # More budget relative to audience = lower marginal score
        reach_curve_penalty = 1.0  # refined once audience sizes known

        scores[ch] = wt * aud_score * efficiency_score * reach_curve_penalty

    total_score = sum(scores.values()) or 1
    split = {}
    for ch in channels:
        split[ch] = round((scores[ch] / total_score) * total_budget, 0)

    # Ensure minimum 5% per channel
    min_budget = total_budget * 0.05
    for ch in split:
        if split[ch] < min_budget:
            split[ch] = min_budget

    # Re-normalise to total
    total_allocated = sum(split.values())
    for ch in split:
        split[ch] = round(split[ch] / total_allocated * total_budget, 0)

    return split

def calculate_channel_kpis(channel, budget_lkr, benchmarks, objective):
    """
    Given a channel budget and historical benchmarks, calculate:
    - KPI type (what to buy on)
    - Buying rate (average from historical data)
    - Target KPI (what we expect to achieve with this budget)
    """
    bench = benchmarks.get(channel, {})

    # Determine KPI type by objective
    objective_kpi_map = {
        "awareness":    ("CPM",  "Impressions"),
        "reach":        ("CPM",  "Impressions"),
        "video":        ("CPV",  "Video Views"),
        "traffic":      ("CPC",  "Clicks"),
        "lead":         ("CPL",  "Leads"),
        "conversion":   ("CPA",  "Conversions"),
        "ecommerce":    ("ROAS", "Revenue"),
        "app":          ("CPC",  "Installs"),
        "engagement":   ("CPE",  "Engagements"),
    }
    obj_lower = objective.lower()
    kpi_type, kpi_metric = next(
        ((v[0], v[1]) for k, v in objective_kpi_map.items() if k in obj_lower),
        ("CPM", "Impressions")
    )

    # Override by channel for search
    if "search" in channel.lower():
        kpi_type, kpi_metric = "CPC", "Clicks"
    elif "youtube" in channel.lower() or "video" in channel.lower():
        kpi_type, kpi_metric = "CPV", "Video Views"
    elif "linkedin" in channel.lower() and "lead" in obj_lower:
        kpi_type, kpi_metric = "CPL", "Leads"

    # Get buying rate from historical data
    rate_key = kpi_type.lower().replace("cpe","cpc")
    historical_rate = bench.get(rate_key)

    if not historical_rate:
        # Fallback rates in LKR if no historical data
        fallback = {"CPM":450,"CPC":95,"CPV":12,"CPA":850,"CPL":1200,"ROAS":None}
        historical_rate = fallback.get(kpi_type, 450)

    buying_rate_lkr = round(historical_rate, 2)

    # Calculate expected KPI from budget
    if kpi_type == "CPM":
        target_kpi = int(budget_lkr / buying_rate_lkr * 1000)
        target_str = f"{target_kpi:,} Impressions"
        rate_str   = f"LKR {buying_rate_lkr:,.0f} CPM"
    elif kpi_type == "CPC":
        target_kpi = int(budget_lkr / buying_rate_lkr)
        target_str = f"{target_kpi:,} Clicks"
        rate_str   = f"LKR {buying_rate_lkr:,.0f} CPC"
    elif kpi_type == "CPV":
        target_kpi = int(budget_lkr / buying_rate_lkr)
        target_str = f"{target_kpi:,} Video Views"
        rate_str   = f"LKR {buying_rate_lkr:,.2f} CPV"
    elif kpi_type == "CPA":
        target_kpi = int(budget_lkr / buying_rate_lkr)
        target_str = f"{target_kpi:,} Conversions"
        rate_str   = f"LKR {buying_rate_lkr:,.0f} CPA"
    elif kpi_type == "CPL":
        target_kpi = int(budget_lkr / buying_rate_lkr)
        target_str = f"{target_kpi:,} Leads"
        rate_str   = f"LKR {buying_rate_lkr:,.0f} CPL"
    elif kpi_type == "ROAS":
        roas = bench.get("roas") or 3.0
        target_str = f"{roas:.1f}x ROAS"
        rate_str   = f"{roas:.1f}x (historical avg)"
    else:
        target_str = "—"
        rate_str   = f"LKR {buying_rate_lkr:,.0f}"

    # CTR benchmark if available
    ctr = bench.get("ctr")
    ctr_str = f"CTR: {ctr:.2f}%" if ctr else ""

    return {
        "kpi_type":    kpi_type,
        "buying_rate": rate_str,
        "target_kpi":  target_str,
        "ctr_bench":   ctr_str,
        "raw_rate":    buying_rate_lkr,
    }

def summarise_dataframe(df):
    lines = [
        f"Rows: {len(df)}, Columns: {len(df.columns)}",
        f"Columns: {', '.join(df.columns.tolist())}",
        "", "Sample data (first 10 rows):",
        df.head(10).to_string(index=False),
        "", "Numeric statistics:",
        df.describe().to_string()
    ]
    col_lower = {c.lower(): c for c in df.columns}
    extras = []
    for metric in ["cpm","cpc","ctr","roas","cpa","cpv","spend","impressions","clicks","conversions","reach"]:
        matches = [c for cl, c in col_lower.items() if metric in cl]
        if matches:
            col = matches[0]
            try:
                vals = pd.to_numeric(df[col], errors="coerce").dropna()
                if len(vals) > 0:
                    extras.append(f"  {col}: mean={vals.mean():.2f}, median={vals.median():.2f}, min={vals.min():.2f}, max={vals.max():.2f}")
            except: pass
    if extras:
        lines += ["", "Key performance benchmarks from historical data:"] + extras
    return "\n".join(lines)

def format_lkr(v):
    try: return f"LKR {float(v):,.0f}"
    except: return "LKR 0"

# ── Chat helpers ──────────────────────────────────────────────────────────────
def render_message(role, content):
    if role == "agent":
        st.markdown(f"""<div class="msg-agent"><div class="avatar-agent">AI</div><div class="bubble-agent">{content.replace(chr(10),'<br>')}</div></div>""", unsafe_allow_html=True)
    else:
        st.markdown(f"""<div class="msg-user"><div class="avatar-user">You</div><div class="bubble-user">{content.replace(chr(10),'<br>')}</div></div>""", unsafe_allow_html=True)

SYSTEM_PROMPT = """You are Orchy, an expert digital media planning agent inside Orchestration-Digital, used by planners in Sri Lanka.

Your job is to collect everything needed to build a complete digital media plan through friendly conversation.

CONVERSATION FLOW — cover in order, 1–2 questions at a time:
1. Campaign name and main objective (Awareness / Traffic / Leads / Conversions / App Installs / Engagement / Video Views)
2. Total budget in LKR and flight dates (start and end)
3. Target audience description (demographics, interests, behaviours)
4. Channels — ask about each: Facebook, Instagram, YouTube, Google Search, Google Display, TikTok, LinkedIn, Programmatic Display
5. For each selected channel, ask for the TARGETABLE AUDIENCE SIZE the planner pulled from the platform's audience estimator (e.g. Meta Audience Insights, Google Reach Planner, TikTok Audience Estimator). Explain they can get this from the platform before planning.
6. For each selected channel, ask what creative assets are available: Static Images, Videos, Carousels, Stories, UGC, Influencer Content — and confirm the objective and placement for each
7. Any additional context or constraints

RULES:
- Be conversational and professional
- Ask 1–2 questions at a time
- Confirm back what you've heard before moving on
- When ALL information is collected, end your message with exactly: [BRIEF_COMPLETE]
- Never generate the media plan yourself
- All budgets in LKR
- When asking for audience sizes, be specific: "How many people does Facebook estimate you can reach with this targeting? You can find this in Meta Audience Insights."

Start by greeting the planner and asking for campaign name and objective."""

def get_agent_response(messages, client_name, brand_name, data_summary):
    client = get_anthropic_client()
    system = SYSTEM_PROMPT + f"\n\nCLIENT: {client_name}\nBRAND: {brand_name}\n\nHISTORICAL DATA SUMMARY:\n{data_summary[:3000]}"
    api_messages = [{"role": m["role"].replace("agent","assistant"), "content": m["content"]} for m in messages if m["role"] != "system"]
    response = client.messages.create(
        model="claude-sonnet-4-6", max_tokens=1000, system=system, messages=api_messages
    )
    return response.content[0].text

def extract_brief_from_conversation(conversation):
    client = get_anthropic_client()
    conv_text = "\n".join([f"{'PLANNER' if m['role']=='user' else 'AGENT'}: {m['content']}" for m in conversation])
    prompt = f"""Extract the campaign brief from this conversation and return ONLY a valid JSON object with these exact fields:
- campaign_name (string)
- objective (string)
- total_budget (number, LKR, no commas)
- start_date (YYYY-MM-DD)
- end_date (YYYY-MM-DD)
- audience (string)
- channels (array of strings — exact channel names mentioned)
- audience_sizes (object: channel name -> integer audience size from platform estimator)
- assets (string — summary of creative assets)
- market (string, default "Sri Lanka")

CONVERSATION:
{conv_text}

Return only valid JSON, no markdown, no explanation."""
    msg = client.messages.create(
        model="claude-sonnet-4-6", max_tokens=800,
        messages=[{"role":"user","content":prompt}]
    )
    try:
        text = msg.content[0].text.strip().replace("```json","").replace("```","").strip()
        return json.loads(text)
    except:
        return {"campaign_name":"Campaign","objective":"","total_budget":0,
                "start_date":str(date.today()),"end_date":str(date.today()),
                "audience":"","channels":[],"audience_sizes":{},"assets":"","market":"Sri Lanka"}

def generate_media_plan(conversation, client_name, brand_name, data_summary, budget_split, channel_kpi_data):
    client = get_anthropic_client()
    conv_text = "\n".join([f"{'PLANNER' if m['role']=='user' else 'AGENT'}: {m['content']}" for m in conversation])

    # Format budget split and KPIs for the prompt
    split_lines = []
    for ch, budget in budget_split.items():
        kpi = channel_kpi_data.get(ch, {})
        split_lines.append(
            f"  {ch}: LKR {budget:,.0f} | KPI Type: {kpi.get('kpi_type','CPM')} | "
            f"Buying Rate: {kpi.get('buying_rate','—')} | Target: {kpi.get('target_kpi','—')}"
        )
    split_summary = "\n".join(split_lines)

    prompt = f"""You are a senior digital media planner. Generate a complete, professional media plan based on the brief and pre-calculated data below.

CLIENT: {client_name}
BRAND: {brand_name}

PRE-CALCULATED BUDGET SPLIT & KPIs (use these exact numbers — do not recalculate):
{split_summary}

PLANNING CONVERSATION:
{conv_text}

HISTORICAL PERFORMANCE DATA:
{data_summary}

Generate the media plan with these sections:

1. EXECUTIVE SUMMARY
   Strategic rationale (3–4 sentences) referencing the brief, objective, and historical performance.

2. BUDGET ALLOCATION RATIONALE
   Explain why the budget was split this way — reference audience sizes, reach efficiency, historical performance, and objective weighting per channel.

3. CHANNEL-BY-CHANNEL PLAN
   For each channel use the pre-calculated numbers above and add:
   - Budget: [from pre-calculated split]
   - Buying Rate: [from pre-calculated data]
   - Target KPI: [from pre-calculated data]
   - CTR Benchmark: [from historical data if available]
   - Ad formats and placements
   - Creative assets mapped to each placement with objectives
   - Targeting approach (use audience size provided)
   - Bidding strategy
   - Weekly pacing

4. CREATIVE ASSET PLAN
   For each asset type: Asset | Format | Platform | Placement | Objective | KPI | Specs

5. REACH & FREQUENCY PROJECTIONS
   Based on audience sizes and budgets — estimate reach % and average frequency per platform.

6. KPI SUMMARY TABLE
   | Channel | Budget (LKR) | KPI Type | Buying Rate | Target KPI | CTR Benchmark |

7. WEEKLY OPTIMISATION ROADMAP

8. RISK FLAGS & MITIGATION

Use LKR throughout. Be precise — every number must come from the pre-calculated data or historical benchmarks."""

    msg = client.messages.create(
        model="claude-sonnet-4-6", max_tokens=5000,
        messages=[{"role":"user","content":prompt}]
    )
    return msg.content[0].text

# ── Excel export ──────────────────────────────────────────────────────────────
def build_excel(brief_summary, plan_text, client_name, brand_name, budget_split, channel_kpi_data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Media Plan"
    navy="1E3A5F"; white="FFFFFF"; light="F0F4FA"; border_col="D0D5E0"

    def cs(row, col, value="", bold=False, bg=None, fg="1A1D23", align="left", size=10, num_fmt=None):
        c = ws.cell(row=row, column=col, value=value)
        c.font = Font(name="Inter", bold=bold, color=fg, size=size)
        if bg: c.fill = PatternFill("solid", fgColor=bg)
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        thin = Side(style="thin", color=border_col)
        c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        if num_fmt: c.number_format = num_fmt
        return c

    def ms(r1,c1,r2,c2,value="",bold=False,bg=None,fg="1A1D23",align="left",size=10):
        ws.merge_cells(start_row=r1,start_column=c1,end_row=r2,end_column=c2)
        c = ws.cell(row=r1,column=c1,value=value)
        c.font = Font(name="Inter",bold=bold,color=fg,size=size)
        if bg: c.fill = PatternFill("solid",fgColor=bg)
        c.alignment = Alignment(horizontal=align,vertical="center",wrap_text=True)
        return c

    for i,w in enumerate([24,30,20,14,18,20,16,18,18,22,12,12,10],1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.row_dimensions[1].height = 40
    ms(1,1,1,13,"ORCHESTRATION-DIGITAL  |  MEDIA PLAN",bold=True,bg=navy,fg=white,align="center",size=14)

    info = brief_summary or {}
    meta = [
        ("Date", datetime.today().strftime("%d %b %Y")),
        ("Client", client_name),
        ("Brand", brand_name),
        ("Campaign Name", info.get("campaign_name","")),
        ("Campaign Start", info.get("start_date","")),
        ("Campaign End",   info.get("end_date","")),
        ("Plan Version", "V1.0"),
    ]
    for i,(label,val) in enumerate(meta,2):
        ws.row_dimensions[i].height = 20
        cs(i,1,label,bold=True,bg=light,fg=navy,align="right")
        ms(i,2,i,6,value=str(val))
        for c in range(7,14): cs(i,c,bg=white)

    hdr_row=10
    ws.row_dimensions[hdr_row].height=40
    hdrs=["Channel","Objective","Target Audience","KPI Type","Buying Rate (LKR)","Target KPI","Spendable (USD)","Spendable (LKR)","Billable (LKR)","Creative Assets","Start Date","End Date","Days"]
    for ci,h in enumerate(hdrs,1):
        cs(hdr_row,ci,h,bold=True,bg=navy,fg=white,align="center",size=9)

    channel_colours={
        "Facebook":"1877F2","Instagram":"E1306C","YouTube":"FF0000",
        "Google Search":"1A8A6E","Google Display":"34A853",
        "TikTok":"010101","LinkedIn":"0A66C2","Programmatic":"7B5EA7",
    }

    usd_rate=320; commission=0.10; ssc_rate=0.025641; vat_rate=0.18; wht_rate=0.163
    channels = list(budget_split.keys())
    start_str = info.get("start_date","")
    end_str   = info.get("end_date","")
    try:
        s=datetime.strptime(start_str,"%Y-%m-%d"); e=datetime.strptime(end_str,"%Y-%m-%d")
        days=(e-s).days
    except: days=0

    current_row=hdr_row+1
    channel_totals={}

    for ch in channels:
        colour=next((v for k,v in channel_colours.items() if k.lower() in ch.lower()),navy)
        ws.row_dimensions[current_row].height=22
        ms(current_row,1,current_row,13,ch.upper(),bold=True,bg=colour,fg=white,align="left",size=10)
        current_row+=1

        sub_lkr = budget_split.get(ch, 0)
        sub_usd = round(sub_lkr/usd_rate, 2)
        billable= round(sub_lkr*1.05, 0)
        row_bg  = light if current_row%2==0 else white
        ws.row_dimensions[current_row].height=22

        kpi_data   = channel_kpi_data.get(ch, {})
        kpi_type   = kpi_data.get("kpi_type","CPM")
        buying_rate= kpi_data.get("buying_rate","—")
        target_kpi = kpi_data.get("target_kpi","—")
        ch_obj     = info.get("objective","")

        data=[ch, ch_obj, info.get("audience",""), kpi_type, buying_rate, target_kpi,
              sub_usd, sub_lkr, billable, info.get("assets",""), start_str, end_str, days]
        for ci,val in enumerate(data,1):
            fmt='#,##0.00' if ci==7 else ('#,##0' if ci in (8,9) else None)
            cs(current_row,ci,val,bg=row_bg,align="center" if ci>3 else "left",num_fmt=fmt)
        channel_totals[ch]=sub_lkr
        current_row+=1

    # Summary rows
    current_row+=1
    total_working=sum(channel_totals.values())
    agency_comm=round(total_working*commission,2)
    sub1=total_working+agency_comm
    ssc=round(sub1*ssc_rate,2)
    sub2=sub1+ssc
    vat=round(sub2*vat_rate,2)
    wht=round(sub2*wht_rate,2)
    total_invest=sub2+vat

    summary=[
        ("Total Working Investment (LKR)",total_working),
        ("Agency Commission (10%)",agency_comm),
        ("Sub Total",sub1),
        ("SSC Levy (2.5641%)",ssc),
        ("Sub Total",sub2),
        ("VAT (18%)",vat),
        ("Withholding Tax (16.3%)",wht),
        ("TOTAL INVESTMENT (LKR)",total_invest),
    ]
    for label,val in summary:
        ws.row_dimensions[current_row].height=22
        is_total="TOTAL INVESTMENT" in label
        is_sub=label.startswith("Sub Total") or label.startswith("Total Working")
        bg=navy if is_total else (light if is_sub else white)
        fg_col=white if is_total else "1A1D23"
        ms(current_row,1,current_row,7,label,bold=is_total or is_sub,bg=bg,fg=fg_col,align="right")
        cs(current_row,8,val,bold=is_total or is_sub,bg=bg,fg=fg_col,align="right",num_fmt='#,##0.00')
        for c in range(9,14): cs(current_row,c,bg=white)
        current_row+=1

    # Sheet 2 — AI Plan
    ws2=wb.create_sheet("AI Plan")
    ws2.column_dimensions["A"].width=120
    ws2.cell(row=1,column=1,value="AI-GENERATED CAMPAIGN PLAN").font=Font(name="Inter",bold=True,size=13,color=navy)
    for i,line in enumerate(plan_text.split("\n"),3):
        c=ws2.cell(row=i,column=1,value=line)
        c.font=Font(name="Inter",size=10)
        c.alignment=Alignment(wrap_text=True)

    # Sheet 3 — KPI Summary
    ws3=wb.create_sheet("KPI Summary")
    ws3.column_dimensions["A"].width=20
    for i,w in enumerate([20,16,18,22,24,20],1):
        ws3.column_dimensions[get_column_letter(i)].width=w
    ws3.merge_cells("A1:F1")
    h=ws3.cell(row=1,column=1,value="KPI SUMMARY BY CHANNEL")
    h.font=Font(name="Inter",bold=True,size=12,color=navy)
    h.fill=PatternFill("solid",fgColor=light)
    kpi_hdrs=["Channel","Budget (LKR)","KPI Type","Buying Rate","Target KPI","Source"]
    for ci,hdr in enumerate(kpi_hdrs,1):
        c=ws3.cell(row=2,column=ci,value=hdr)
        c.font=Font(name="Inter",bold=True,color=white,size=9)
        c.fill=PatternFill("solid",fgColor=navy)
        c.alignment=Alignment(horizontal="center",vertical="center")
    for i,(ch,budget) in enumerate(budget_split.items(),3):
        kpi=channel_kpi_data.get(ch,{})
        row_data=[ch,budget,kpi.get("kpi_type","—"),kpi.get("buying_rate","—"),kpi.get("target_kpi","—"),"Historical Data Avg"]
        for ci,val in enumerate(row_data,1):
            c=ws3.cell(row=i,column=ci,value=val)
            c.font=Font(name="Inter",size=10)
            c.fill=PatternFill("solid",fgColor=("F0F4FA" if i%2==0 else "FFFFFF"))
            if ci==2:
                c.number_format='#,##0'
            c.alignment=Alignment(horizontal="center" if ci>1 else "left",vertical="center")

    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()

# ── Session defaults ──────────────────────────────────────────────────────────
for key,default in [("step",1),("brief",{}),("selected_client",None),
                    ("selected_brand",None),("combined_df",None),
                    ("generated_plan",None),("chat_messages",[]),
                    ("brief_summary",{}),("budget_split",{}),("channel_kpi_data",{})]:
    if key not in st.session_state:
        st.session_state[key]=default

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='padding:8px 0 24px 0;'>
      <div style='font-family:Plus Jakarta Sans,sans-serif;font-size:1.25rem;font-weight:800;color:#1e3a5f;'>🎯 Orchestration</div>
      <div style='font-size:0.75rem;color:#8a93a8;margin-top:2px;'>Digital Planning Platform</div>
    </div>""", unsafe_allow_html=True)
    nav=st.radio("Navigation",[
        "📊 New Campaign Plan","📁 Saved Plans",
        "📈 Data Explorer","⚙️ Settings"
    ],label_visibility="collapsed")
    st.markdown("---")
    st.markdown("<div style='font-size:0.72rem;color:#8a93a8;'>POC Version 1.0 · Phase 1</div>",unsafe_allow_html=True)

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
if nav=="📊 New Campaign Plan":
    sb=get_supabase()

    steps=["1 · Client & Brand","2 · Historical Data","3 · Plan with Orchy","4 · Media Plan"]
    cols=st.columns(4)
    for i,(col,label) in enumerate(zip(cols,steps),1):
        active=st.session_state["step"]==i
        done=st.session_state["step"]>i
        bg="#1e3a5f" if active else ("#1a8a6e" if done else "#e8eaf0")
        fg="#ffffff" if (active or done) else "#8a93a8"
        icon="✓" if done else str(i)
        col.markdown(f"""<div style='background:{bg};color:{fg};border-radius:10px;padding:12px 16px;text-align:center;font-weight:600;font-size:0.82rem;'>{icon} · {label.split("·")[1].strip()}</div>""",unsafe_allow_html=True)
    st.markdown("<br>",unsafe_allow_html=True)

    # STEP 1
    if st.session_state["step"]==1:
        st.markdown('<div class="section-header">Select Client & Brand</div>',unsafe_allow_html=True)
        clients=get_clients_list(sb)
        client_names=[c["client_name"] for c in clients]
        col_a,col_b=st.columns(2)
        with col_a:
            st.markdown("**Client**")
            client_choice=st.selectbox("Client",["— Select existing —","➕ Add new client"]+client_names,label_visibility="collapsed")
            if client_choice=="➕ Add new client":
                new_client=st.text_input("New client name",placeholder="e.g. Unilever Lanka")
                if st.button("Create Client") and new_client:
                    r=create_client_record(sb,new_client)
                    if r: st.success(f"✅ '{new_client}' created!"); st.rerun()
            elif client_choice!="— Select existing —":
                st.session_state["selected_client"]=next((c for c in clients if c["client_name"]==client_choice),None)
        with col_b:
            if st.session_state.get("selected_client"):
                st.markdown("**Brand**")
                brands=get_brands_for_client(sb,st.session_state["selected_client"]["id"])
                brand_names=[b["brand_name"] for b in brands]
                brand_choice=st.selectbox("Brand",["— Select existing —","➕ Add new brand"]+brand_names,label_visibility="collapsed")
                if brand_choice=="➕ Add new brand":
                    new_brand=st.text_input("New brand name",placeholder="e.g. Sunlight")
                    if st.button("Create Brand") and new_brand:
                        r=create_brand_record(sb,st.session_state["selected_client"]["id"],new_brand)
                        if r: st.success(f"✅ '{new_brand}' created!"); st.rerun()
                elif brand_choice!="— Select existing —":
                    st.session_state["selected_brand"]=next((b for b in brands if b["brand_name"]==brand_choice),None)
            else:
                st.info("Select a client first.")
        if st.session_state.get("selected_client") and st.session_state.get("selected_brand"):
            st.markdown("---")
            st.markdown('<div class="green-btn">',unsafe_allow_html=True)
            if st.button("Continue to Historical Data →",use_container_width=True):
                st.session_state["step"]=2
                st.session_state["chat_messages"]=[]
                st.session_state["generated_plan"]=None
                st.session_state["budget_split"]={}
                st.session_state["channel_kpi_data"]={}
                st.rerun()
            st.markdown('</div>',unsafe_allow_html=True)

    # STEP 2
    elif st.session_state["step"]==2:
        brand_id=st.session_state["selected_brand"]["id"]
        brand_name=st.session_state["selected_brand"]["brand_name"]
        client_name=st.session_state["selected_client"]["client_name"]
        st.markdown(f'<div class="section-header">Historical Data — {client_name} · {brand_name}</div>',unsafe_allow_html=True)
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
                    if cc.button("🗑",key=f"del_{f['id']}"):
                        delete_brand_data_file(sb,f["id"]); st.rerun()
                    try: dfs.append(pd.read_json(io.StringIO(f["data_json"])))
                    except: pass
            if dfs:
                st.session_state["combined_df"]=pd.concat(dfs,ignore_index=True)
        else:
            st.info(f"No historical data saved for **{brand_name}** yet.")
        st.markdown('<div class="section-header">Upload New Files</div>',unsafe_allow_html=True)
        uploaded_files=st.file_uploader("Drop files here",type=["csv","xlsx","xls"],accept_multiple_files=True)
        if uploaded_files:
            new_dfs=[]
            for f in uploaded_files:
                df=parse_uploaded_file(f)
                if df is not None: new_dfs.append((f.name,df))
            if new_dfs:
                if st.button(f"💾 Save {len(new_dfs)} file(s) to {brand_name}",use_container_width=True):
                    for fname,df in new_dfs:
                        save_brand_data(sb,brand_id,fname,df.to_json(),len(df))
                    st.success("✅ Saved!"); st.rerun()
                existing=st.session_state.get("combined_df")
                new_combined=pd.concat([d for _,d in new_dfs],ignore_index=True)
                st.session_state["combined_df"]=pd.concat([existing,new_combined],ignore_index=True) if existing is not None else new_combined
                with st.expander("Preview"):
                    st.dataframe(new_combined.head(20),use_container_width=True)
        st.markdown("---")
        c1,c2=st.columns(2)
        with c1:
            if st.button("← Back",use_container_width=True):
                st.session_state["step"]=1; st.rerun()
        with c2:
            st.markdown('<div class="green-btn">',unsafe_allow_html=True)
            if st.button("Continue to Planning Chat →",use_container_width=True):
                if st.session_state.get("combined_df") is None:
                    st.error("Please upload at least one data file.")
                else:
                    st.session_state["step"]=3; st.rerun()
            st.markdown('</div>',unsafe_allow_html=True)

    # STEP 3 — CHAT
    elif st.session_state["step"]==3:
        client_name=st.session_state["selected_client"]["client_name"]
        brand_name=st.session_state["selected_brand"]["brand_name"]
        data_summary=summarise_dataframe(st.session_state["combined_df"])

        st.markdown(f'<div class="section-header">Plan with Orchy — {client_name} · {brand_name}</div>',unsafe_allow_html=True)
        st.caption("Orchy will guide you through the brief. Answer the questions to build your campaign plan.")

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

        st.markdown('<div class="chat-container"><div class="chat-header"><p class="chat-header-title">🤖 Orchy — Planning Agent</p><p class="chat-header-sub">AI-powered digital planning assistant</p></div><div class="chat-messages">',unsafe_allow_html=True)
        for msg in st.session_state["chat_messages"]:
            render_message(msg["role"],msg["content"].replace("[BRIEF_COMPLETE]","").strip())
        st.markdown('</div></div>',unsafe_allow_html=True)

        last_agent=next((m["content"] for m in reversed(st.session_state["chat_messages"]) if m["role"]=="agent"),"")
        brief_auto_complete="[BRIEF_COMPLETE]" in last_agent
        num_user_msgs=len([m for m in st.session_state["chat_messages"] if m["role"]=="user"])

        with st.form("chat_form",clear_on_submit=True):
            col_inp,col_btn=st.columns([5,1])
            user_input=col_inp.text_input("Your message",placeholder="Type your answer here…",label_visibility="collapsed")
            submitted=col_btn.form_submit_button("Send →")

        if submitted and user_input.strip():
            st.session_state["chat_messages"].append({"role":"user","content":user_input.strip()})
            with st.spinner("Orchy is thinking…"):
                reply=get_agent_response(st.session_state["chat_messages"],client_name,brand_name,data_summary)
            st.session_state["chat_messages"].append({"role":"agent","content":reply})
            st.rerun()

        if num_user_msgs>=3:
            st.markdown("---")
            if brief_auto_complete:
                st.success("✅ Orchy has all the information needed. Ready to generate your media plan!")
            else:
                st.info("💡 Ready to generate? Click below — or keep chatting to add more detail.")
            st.markdown('<div class="green-btn">',unsafe_allow_html=True)
            if st.button("🚀 Generate Media Plan",use_container_width=True):
                with st.spinner("🤖 Calculating budget split, buying rates and KPI targets from your historical data…"):
                    try:
                        # 1. Extract brief
                        brief_summary=extract_brief_from_conversation(st.session_state["chat_messages"])
                        st.session_state["brief_summary"]=brief_summary

                        # 2. Extract benchmarks from historical data
                        benchmarks=extract_channel_benchmarks(st.session_state["combined_df"])

                        # 3. Get platform settings
                        platform_settings=get_platform_settings(sb)
                        ps_dict={p["platform_name"]:p for p in platform_settings}

                        # 4. Calculate optimised budget split
                        channels=brief_summary.get("channels",[])
                        audience_sizes=brief_summary.get("audience_sizes",{})
                        total_budget=float(brief_summary.get("total_budget",0))
                        objective=brief_summary.get("objective","Brand Awareness")

                        budget_split=calculate_budget_split(
                            channels,total_budget,objective,
                            audience_sizes,benchmarks,ps_dict
                        )
                        st.session_state["budget_split"]=budget_split

                        # 5. Calculate KPIs per channel
                        channel_kpi_data={}
                        for ch,budget in budget_split.items():
                            channel_kpi_data[ch]=calculate_channel_kpis(ch,budget,benchmarks,objective)
                        st.session_state["channel_kpi_data"]=channel_kpi_data

                        # 6. Generate plan text
                        plan_text=generate_media_plan(
                            st.session_state["chat_messages"],
                            client_name,brand_name,data_summary,
                            budget_split,channel_kpi_data
                        )
                        st.session_state["generated_plan"]=plan_text
                        st.session_state["step"]=4
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error generating plan: {e}")
            st.markdown('</div>',unsafe_allow_html=True)

        st.markdown("---")
        if st.button("← Back to Data"):
            st.session_state["step"]=2; st.rerun()

    # STEP 4 — PLAN
    elif st.session_state["step"]==4:
        client_name=st.session_state["selected_client"]["client_name"]
        brand_name=st.session_state["selected_brand"]["brand_name"]
        plan_text=st.session_state["generated_plan"]
        brief_summary=st.session_state.get("brief_summary",{})
        budget_split=st.session_state.get("budget_split",{})
        channel_kpi_data=st.session_state.get("channel_kpi_data",{})

        st.markdown(f'<div class="section-header">{client_name} · {brand_name} · {brief_summary.get("campaign_name","Media Plan")}</div>',unsafe_allow_html=True)

        total_budget=float(brief_summary.get("total_budget",0))
        try:
            s=datetime.strptime(brief_summary.get("start_date",""),"%Y-%m-%d")
            e=datetime.strptime(brief_summary.get("end_date",""),"%Y-%m-%d")
            days=(e-s).days
        except: days="—"

        c1,c2,c3,c4=st.columns(4)
        with c1: st.markdown(f'<div class="metric-card"><div class="metric-value">{format_lkr(total_budget)}</div><div class="metric-label">Total Budget</div></div>',unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-card"><div class="metric-value">{days} days</div><div class="metric-label">Duration</div></div>',unsafe_allow_html=True)
        with c3: st.markdown(f'<div class="metric-card"><div class="metric-value">{len(budget_split)}</div><div class="metric-label">Channels</div></div>',unsafe_allow_html=True)
        with c4:
            obj=brief_summary.get("objective","—")
            st.markdown(f'<div class="metric-card"><div class="metric-value">{obj.split()[0] if obj else "—"}</div><div class="metric-label">Objective</div></div>',unsafe_allow_html=True)

        # Budget split table
        if budget_split:
            st.markdown('<div class="section-header">Budget Split & KPI Targets</div>',unsafe_allow_html=True)
            rows=[]
            for ch,budget in budget_split.items():
                kpi=channel_kpi_data.get(ch,{})
                rows.append({
                    "Channel":ch,
                    "Budget (LKR)":f"LKR {budget:,.0f}",
                    "% of Total":f"{budget/total_budget*100:.1f}%" if total_budget>0 else "—",
                    "KPI Type":kpi.get("kpi_type","—"),
                    "Buying Rate":kpi.get("buying_rate","—"),
                    "Target KPI":kpi.get("target_kpi","—"),
                })
            st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)

        st.markdown('<div class="section-header">Full Media Plan</div>',unsafe_allow_html=True)
        st.markdown(f'<div class="plan-card">{plan_text.replace(chr(10),"<br>")}</div>',unsafe_allow_html=True)

        st.markdown("---")
        ca,cb,cc,cd=st.columns(4)
        with ca:
            st.download_button("📥 Download TXT",data=plan_text,
                file_name=f"{brief_summary.get('campaign_name','plan').replace(' ','_')}_plan.txt",
                mime="text/plain",use_container_width=True)
        with cb:
            excel_bytes=build_excel(brief_summary,plan_text,client_name,brand_name,budget_split,channel_kpi_data)
            st.download_button("📊 Download Excel",data=excel_bytes,
                file_name=f"{brief_summary.get('campaign_name','plan').replace(' ','_')}_MediaPlan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with cc:
            if st.button("💾 Save to Library",use_container_width=True):
                save_plan(sb,{
                    "brand_id":      st.session_state["selected_brand"]["id"],
                    "client_name":   client_name,
                    "brand_name":    brand_name,
                    "campaign_name": brief_summary.get("campaign_name",""),
                    "objective":     brief_summary.get("objective",""),
                    "total_budget":  total_budget,
                    "start_date":    brief_summary.get("start_date",""),
                    "end_date":      brief_summary.get("end_date",""),
                    "channels":      json.dumps(list(budget_split.keys())),
                    "market":        brief_summary.get("market","Sri Lanka"),
                    "kpi_focus":     "",
                    "plan_text":     plan_text,
                    "created_at":    datetime.utcnow().isoformat()
                })
                st.success("Saved!")
        with cd:
            st.markdown('<div class="green-btn">',unsafe_allow_html=True)
            if st.button("➕ New Plan",use_container_width=True):
                for k in ["step","generated_plan","chat_messages","brief_summary","budget_split","channel_kpi_data"]:
                    st.session_state[k]=1 if k=="step" else ([] if k=="chat_messages" else {})
                st.rerun()
            st.markdown('</div>',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SAVED PLANS
# ══════════════════════════════════════════════════════════════════════════════
elif nav=="📁 Saved Plans":
    sb=get_supabase()
    st.markdown('<div class="section-header">Saved Campaign Plans</div>',unsafe_allow_html=True)
    try:
        plans=load_saved_plans(sb)
        if not plans:
            st.info("No saved plans yet.")
        else:
            for plan in plans:
                budget=float(plan.get("total_budget",0) or 0)
                label=f"📋 {plan.get('client_name','—')} · {plan.get('brand_name','—')} · {plan.get('campaign_name','Unnamed')} — {format_lkr(budget)}"
                with st.expander(label):
                    c1,c2,c3=st.columns(3)
                    c1.metric("Objective",plan.get("objective","—"))
                    c2.metric("Budget",format_lkr(budget))
                    c3.metric("Market",plan.get("market","—"))
                    st.markdown(f"**Dates:** {plan.get('start_date','')} → {plan.get('end_date','')}")
                    st.markdown("---")
                    st.markdown(plan.get("plan_text",""))
    except Exception as e:
        st.error(f"Could not load plans: {e}")

# ══════════════════════════════════════════════════════════════════════════════
# DATA EXPLORER
# ══════════════════════════════════════════════════════════════════════════════
elif nav=="📈 Data Explorer":
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
elif nav=="⚙️ Settings":
    sb=get_supabase()
    st.markdown('<div class="section-header">Platform Settings</div>',unsafe_allow_html=True)
    st.caption("Configure platform efficiency benchmarks. These are used to optimise budget allocation across channels. Audience sizes are captured per campaign during the planning chat.")

    platforms=[
        {"name":"Facebook",       "default_freq_awareness":3,"default_freq_conversion":7,"reach_curve_inflection":0.4},
        {"name":"Instagram",      "default_freq_awareness":3,"default_freq_conversion":6,"reach_curve_inflection":0.35},
        {"name":"YouTube",        "default_freq_awareness":4,"default_freq_conversion":5,"reach_curve_inflection":0.45},
        {"name":"Google Search",  "default_freq_awareness":1,"default_freq_conversion":3,"reach_curve_inflection":0.6},
        {"name":"Google Display", "default_freq_awareness":5,"default_freq_conversion":8,"reach_curve_inflection":0.3},
        {"name":"TikTok",         "default_freq_awareness":3,"default_freq_conversion":5,"reach_curve_inflection":0.38},
        {"name":"LinkedIn",       "default_freq_awareness":4,"default_freq_conversion":6,"reach_curve_inflection":0.5},
        {"name":"Programmatic",   "default_freq_awareness":6,"default_freq_conversion":9,"reach_curve_inflection":0.3},
    ]

    existing_settings=get_platform_settings(sb)
    existing_dict={p["platform_name"]:p for p in existing_settings}

    st.markdown('<div class="section-header">Frequency Caps & Reach Curve Parameters</div>',unsafe_allow_html=True)
    st.caption("Frequency cap = max times a user sees your ad. Reach curve inflection = point (as % of audience) where diminishing returns begin.")

    for p in platforms:
        pname=p["name"]
        saved=existing_dict.get(pname,{})
        with st.expander(f"⚙️ {pname}"):
            col1,col2,col3=st.columns(3)
            freq_aw=col1.number_input(f"Freq Cap — Awareness",min_value=1,max_value=20,
                value=int(saved.get("freq_cap_awareness",p["default_freq_awareness"])),key=f"fa_{pname}")
            freq_cv=col2.number_input(f"Freq Cap — Conversion",min_value=1,max_value=20,
                value=int(saved.get("freq_cap_conversion",p["default_freq_conversion"])),key=f"fc_{pname}")
            reach_inf=col3.number_input(f"Reach Curve Inflection (%)",min_value=0.1,max_value=1.0,step=0.05,
                value=float(saved.get("reach_curve_inflection",p["reach_curve_inflection"])),key=f"ri_{pname}")
            notes=st.text_input("Notes",value=saved.get("notes",""),key=f"n_{pname}",
                placeholder="e.g. Strong for video in LK, peak hours 7-10pm")
            if st.button(f"Save {pname} Settings",key=f"save_{pname}"):
                save_platform_setting(sb,pname,{
                    "freq_cap_awareness":freq_aw,
                    "freq_cap_conversion":freq_cv,
                    "reach_curve_inflection":reach_inf,
                    "notes":notes
                })
                st.success(f"✅ {pname} settings saved!")

    st.markdown('<div class="section-header">About Budget Optimisation</div>',unsafe_allow_html=True)
    st.markdown("""
    <div class="settings-card">
    <p style='font-size:0.88rem; color:#4a5168; line-height:1.7;'>
    Budget is allocated using a multi-factor model:<br><br>
    <b>1. Objective weighting</b> — channels are scored by how well they match the campaign objective (e.g. YouTube scores higher for Video Views, Google Search for Traffic).<br>
    <b>2. Audience size</b> — captured during planning chat from each platform's estimator (Meta Audience Insights, Google Reach Planner, TikTok Audience Estimator). Larger targetable audiences allow more budget headroom before diminishing returns.<br>
    <b>3. Historical efficiency</b> — lower historical CPM/CPC = more efficient channel = higher budget score.<br>
    <b>4. Reach curve</b> — the inflection point above controls when diminishing returns kick in. Budget beyond this point has lower marginal reach gain.<br><br>
    <i>In Phase 2, audience sizes will be pulled directly from platform APIs.</i>
    </p>
    </div>""", unsafe_allow_html=True)