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

  /* Chat styles */
  .chat-container { background:#fff; border:1px solid #e8eaf0; border-radius:14px; padding:0; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.05); margin-bottom:16px; }
  .chat-header { background:linear-gradient(135deg,#1e3a5f,#2d5a9b); padding:16px 20px; }
  .chat-header-title { color:#fff; font-family:'Plus Jakarta Sans',sans-serif; font-weight:700; font-size:0.95rem; margin:0; }
  .chat-header-sub { color:rgba(255,255,255,0.7); font-size:0.78rem; margin:2px 0 0 0; }
  .chat-messages { padding:20px; max-height:500px; overflow-y:auto; }

  .msg-agent { display:flex; gap:12px; margin-bottom:16px; align-items:flex-start; }
  .msg-user  { display:flex; gap:12px; margin-bottom:16px; align-items:flex-start; flex-direction:row-reverse; }

  .avatar-agent { width:32px; height:32px; background:linear-gradient(135deg,#1e3a5f,#2d5a9b); border-radius:50%; display:flex; align-items:center; justify-content:center; color:#fff; font-size:0.75rem; font-weight:700; flex-shrink:0; }
  .avatar-user  { width:32px; height:32px; background:linear-gradient(135deg,#1a8a6e,#22b894); border-radius:50%; display:flex; align-items:center; justify-content:center; color:#fff; font-size:0.75rem; font-weight:700; flex-shrink:0; }

  .bubble-agent { background:#f0f4fa; border-radius:0 12px 12px 12px; padding:12px 16px; max-width:80%; font-size:0.88rem; line-height:1.6; color:#1a1d23; }
  .bubble-user  { background:linear-gradient(135deg,#1e3a5f,#2d5a9b); border-radius:12px 0 12px 12px; padding:12px 16px; max-width:80%; font-size:0.88rem; line-height:1.6; color:#fff; }

  .plan-card { background:#fff; border:1px solid #e8eaf0; border-radius:14px; padding:28px 32px; margin-top:16px; line-height:1.8; color:#2d3341; box-shadow:0 1px 4px rgba(0,0,0,0.05); font-size:0.92rem; }

  .stButton > button { background:linear-gradient(135deg,#1e3a5f,#2d5a9b) !important; color:white !important; border:none !important; border-radius:8px !important; font-weight:600 !important; font-size:0.9rem !important; padding:10px 24px !important; }
  .stButton > button:hover { opacity:0.88 !important; }
  .green-btn .stButton > button { background:linear-gradient(135deg,#1a8a6e,#22b894) !important; }

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

# ── Clients ───────────────────────────────────────────────────────────────────
@st.cache_resource
def get_supabase() -> Client:
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

def get_anthropic_client():
    return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])

# ── Supabase helpers ──────────────────────────────────────────────────────────
def get_clients_list(sb):
    try:
        return sb.table("clients").select("*").order("client_name").execute().data or []
    except: return []

def get_brands_for_client(sb, client_id):
    try:
        return sb.table("brands").select("*").eq("client_id", client_id).order("brand_name").execute().data or []
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
    try:
        return sb.table("brand_data").select("*").eq("brand_id", brand_id).order("uploaded_at", desc=True).execute().data or []
    except: return []

def save_brand_data(sb, brand_id, file_name, data_json, row_count):
    try:
        sb.table("brand_data").insert({"brand_id": brand_id, "file_name": file_name, "data_json": data_json, "row_count": row_count, "uploaded_at": datetime.utcnow().isoformat()}).execute()
        return True
    except Exception as e:
        st.error(f"Error saving data: {e}"); return False

def delete_brand_data_file(sb, file_id):
    try:
        sb.table("brand_data").delete().eq("id", file_id).execute(); return True
    except: return False

def save_plan(sb, plan_data):
    try:
        sb.table("campaign_plans").insert(plan_data).execute(); return True
    except Exception as e:
        st.warning(f"Could not save: {e}"); return False

def load_saved_plans(sb):
    try:
        return sb.table("campaign_plans").select("*").order("created_at", desc=True).limit(30).execute().data or []
    except: return []

# ── Data helpers ──────────────────────────────────────────────────────────────
def parse_uploaded_file(f):
    n = f.name.lower()
    if n.endswith(".csv"): return pd.read_csv(f)
    elif n.endswith((".xlsx",".xls")): return pd.read_excel(f)
    return None

def summarise_dataframe(df):
    lines = [
        f"Rows: {len(df)}, Columns: {len(df.columns)}",
        f"Columns: {', '.join(df.columns.tolist())}",
        "", "Sample data (first 10 rows):",
        df.head(10).to_string(index=False),
        "", "Numeric statistics:",
        df.describe().to_string()
    ]
    # Try to extract key metrics for benchmarking
    num_cols = df.select_dtypes(include='number').columns.tolist()
    col_lower = {c.lower(): c for c in df.columns}
    extras = []
    for metric in ["cpm","cpc","ctr","roas","cpa","cpv","spend","impressions","clicks","conversions"]:
        matches = [c for cl, c in col_lower.items() if metric in cl]
        if matches:
            col = matches[0]
            try:
                extras.append(f"  {col}: mean={df[col].mean():.2f}, median={df[col].median():.2f}, min={df[col].min():.2f}, max={df[col].max():.2f}")
            except: pass
    if extras:
        lines += ["", "Key performance benchmarks from historical data:"] + extras
    return "\n".join(lines)

def format_lkr(v):
    return f"LKR {v:,.0f}"

# ── Chat helpers ──────────────────────────────────────────────────────────────
def render_message(role, content):
    if role == "agent":
        st.markdown(f"""
        <div class="msg-agent">
          <div class="avatar-agent">AI</div>
          <div class="bubble-agent">{content.replace(chr(10),'<br>')}</div>
        </div>""", unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="msg-user">
          <div class="avatar-user">You</div>
          <div class="bubble-user">{content.replace(chr(10),'<br>')}</div>
        </div>""", unsafe_allow_html=True)

SYSTEM_PROMPT = """You are an expert digital media planning agent called Orchy, working inside Orchestration-Digital — a professional media planning platform used by digital planners in Sri Lanka.

Your job is to conduct a friendly, professional conversation to collect everything needed to build a complete digital media plan. You already know the client name, brand name, and have access to historical campaign performance data.

CONVERSATION FLOW — ask about these topics in order, one or two at a time:
1. Campaign name and main objective (Awareness / Traffic / Leads / Conversions / App Installs / Engagement / Video Views)
2. Total budget in LKR and campaign flight dates (start and end)
3. Target audience description (demographics, interests, behaviours)
4. Channel selection — ask about each separately: Facebook, Instagram, YouTube, Google Search, Google Display, TikTok, LinkedIn, Programmatic Display
5. For each selected channel, ask what creative assets are available:
   - Static Images, Videos, Carousels, Stories, UGC (User Generated Content), Influencer Content
   - For each asset type, confirm the objective and placement
6. Any additional context, special requirements, or constraints

RULES:
- Be conversational and friendly but professional
- Ask 1-2 questions at a time, never overwhelm
- Confirm back what you've heard before moving on
- When you have collected ALL information, end your message with exactly: [BRIEF_COMPLETE]
- Never generate the media plan yourself — just collect the brief
- All budgets are in LKR (Sri Lankan Rupees)
- Use your knowledge of Sri Lanka digital market norms

Start by greeting the planner and asking for the campaign name and objective."""

def get_agent_response(messages, client_name, brand_name, data_summary):
    client = get_anthropic_client()
    system = SYSTEM_PROMPT + f"\n\nCLIENT: {client_name}\nBRAND: {brand_name}\n\nHISTORICAL DATA SUMMARY:\n{data_summary[:3000]}"
    api_messages = [{"role": m["role"].replace("agent","assistant"), "content": m["content"]} for m in messages if m["role"] != "system"]
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1000,
        system=system,
        messages=api_messages
    )
    return response.content[0].text

def generate_media_plan(conversation, client_name, brand_name, data_summary):
    client = get_anthropic_client()
    conv_text = "\n".join([f"{'PLANNER' if m['role']=='user' else 'AGENT'}: {m['content']}" for m in conversation])

    prompt = f"""You are a senior digital media planner. Based on the planning conversation below and the historical campaign data, generate a comprehensive media plan.

CLIENT: {client_name}
BRAND: {brand_name}

PLANNING CONVERSATION:
{conv_text}

HISTORICAL PERFORMANCE DATA:
{data_summary}

Generate a complete media plan with these sections:

1. EXECUTIVE SUMMARY
   Strategic rationale (3–4 sentences) referencing the brief and historical performance.

2. CHANNEL MIX & BUDGET ALLOCATION
   For each channel:
   - Budget in LKR (amount and % of total)
   - Rationale based on historical data
   - Expected reach or impressions
   - Flight dates

3. CHANNEL-BY-CHANNEL TACTICS
   For each channel include:
   - Recommended buying rate type (CPM / CPC / CPV / CPA / CPL)
   - TARGET BUYING RATE in LKR — calculated from historical data benchmarks (give a specific range e.g. "CPM: LKR 380–450")
   - Primary KPI with target value based on historical performance (e.g. "Target CTR: 1.2–1.8% based on past campaigns")
   - Secondary KPI targets with benchmark ranges
   - Ad formats and placements
   - Creative assets mapped to each placement (Static / Video / Carousel / Stories / UGC / Influencer)
   - Objective per creative asset type
   - Targeting approach
   - Bidding strategy

4. CREATIVE ASSET PLAN
   Table format for each channel:
   - Asset type | Format | Placement | Objective | KPI | Recommended specs

5. KPI TARGETS & BENCHMARKS
   Summary table of all KPIs with:
   - Historical benchmark (from data)
   - Target for this campaign
   - Minimum acceptable threshold

6. WEEKLY OPTIMISATION ROADMAP
   Week-by-week actions.

7. RISK FLAGS
   Risks from historical data and mitigation strategies.

Use LKR for all budget figures. Ground every number in the historical data provided. Be specific and professional."""

    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=5000,
        messages=[{"role": "user", "content": prompt}]
    )
    return msg.content[0].text

# ── Excel export ──────────────────────────────────────────────────────────────
def build_excel(brief_summary, plan_text, client_name, brand_name):
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

    for i,w in enumerate([22,36,18,16,14,16,16,18,18,36,14,14,10],1):
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
        ms(i,2,i,5,value=str(val))
        for c in range(6,14): cs(i,c,bg=white)

    hdr_row=10
    ws.row_dimensions[hdr_row].height=36
    hdrs=["Channel / Placement","Campaign Objective","Audience","KPI Type","Buying Rate (LKR)","Target KPI","Spendable (USD)","Spendable (LKR)","Billable (LKR)","Creative Assets","Start Date","End Date","Days"]
    for ci,h in enumerate(hdrs,1):
        cs(hdr_row,ci,h,bold=True,bg=navy,fg=white,align="center",size=9)

    channel_colours = {
        "Facebook":           "1877F2",
        "Instagram":          "E1306C",
        "YouTube":            "FF0000",
        "Google Search":      "1A8A6E",
        "Google Display":     "34A853",
        "TikTok":             "010101",
        "LinkedIn":           "0A66C2",
        "Programmatic":       "7B5EA7",
    }

    channels = info.get("channels", list(channel_colours.keys()))
    total_lkr = info.get("total_budget", 0)
    start_str = info.get("start_date","")
    end_str   = info.get("end_date","")
    try:
        s = datetime.strptime(start_str,"%Y-%m-%d"); e = datetime.strptime(end_str,"%Y-%m-%d")
        days = (e-s).days
    except: days = 0

    n = len(channels) if channels else 1
    per_ch = round(total_lkr/n, 0)
    usd_rate=320; commission=0.10; ssc_rate=0.025641; vat_rate=0.18; wht_rate=0.163
    current_row = hdr_row+1
    channel_totals = {}

    for ch in channels:
        colour = next((v for k,v in channel_colours.items() if k.lower() in ch.lower()), navy)
        ws.row_dimensions[current_row].height=22
        ms(current_row,1,current_row,13,ch.upper(),bold=True,bg=colour,fg=white,align="left",size=10)
        current_row+=1
        sub_lkr = per_ch
        sub_usd = round(sub_lkr/usd_rate,2)
        billable= round(sub_lkr*1.05,0)
        row_bg = light if current_row%2==0 else white
        ws.row_dimensions[current_row].height=20
        data=[ch, info.get("objective",""), info.get("audience",""), "CPM","","",sub_usd,sub_lkr,billable,info.get("assets",""),start_str,end_str,days]
        for ci,val in enumerate(data,1):
            fmt='#,##0.00' if ci==7 else ('#,##0' if ci in (8,9) else None)
            cs(current_row,ci,val,bg=row_bg,align="center" if ci>5 else "left",num_fmt=fmt)
        channel_totals[ch] = sub_lkr
        current_row+=1

    current_row+=1
    total_working = sum(channel_totals.values())
    agency_comm   = round(total_working*commission,2)
    sub1 = total_working+agency_comm
    ssc  = round(sub1*ssc_rate,2)
    sub2 = sub1+ssc
    vat  = round(sub2*vat_rate,2)
    wht  = round(sub2*wht_rate,2)
    total_invest = sub2+vat

    summary=[("Total Working Investment (LKR)",total_working),("Agency Commission (10%)",agency_comm),
             ("Sub Total",sub1),("SSC Levy (2.5641%)",ssc),("Sub Total",sub2),
             ("VAT (18%)",vat),("Withholding Tax (16.3%)",wht),("TOTAL INVESTMENT (LKR)",total_invest)]
    for label,val in summary:
        ws.row_dimensions[current_row].height=22
        is_total="TOTAL INVESTMENT" in label
        is_sub=label.startswith("Sub Total") or label.startswith("Total Working")
        bg=navy if is_total else (light if is_sub else white)
        fg_col=white if is_total else "1A1D23"
        ms(current_row,1,current_row,8,label,bold=is_total or is_sub,bg=bg,fg=fg_col,align="right")
        cs(current_row,9,val,bold=is_total or is_sub,bg=bg,fg=fg_col,align="right",num_fmt='#,##0.00')
        for c in range(10,14): cs(current_row,c,bg=white)
        current_row+=1

    ws2=wb.create_sheet("AI Plan")
    ws2.column_dimensions["A"].width=120
    ws2.cell(row=1,column=1,value="AI-GENERATED CAMPAIGN PLAN").font=Font(name="Inter",bold=True,size=13,color=navy)
    for i,line in enumerate(plan_text.split("\n"),3):
        c=ws2.cell(row=i,column=1,value=line)
        c.font=Font(name="Inter",size=10)
        c.alignment=Alignment(wrap_text=True)

    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()

# ── Extract brief summary from conversation ───────────────────────────────────
def extract_brief_from_conversation(conversation, client_name, brand_name, data_summary):
    client = get_anthropic_client()
    conv_text = "\n".join([f"{'PLANNER' if m['role']=='user' else 'AGENT'}: {m['content']}" for m in conversation])
    prompt = f"""Extract the campaign brief details from this conversation and return ONLY a JSON object with these fields:
campaign_name, objective, total_budget (number in LKR), start_date (YYYY-MM-DD), end_date (YYYY-MM-DD),
audience, channels (array of strings), assets (string summary of creative assets), market.

CONVERSATION:
{conv_text}

Return only valid JSON, no other text."""
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=500,
        messages=[{"role":"user","content":prompt}]
    )
    try:
        text = msg.content[0].text.strip()
        text = text.replace("```json","").replace("```","").strip()
        return json.loads(text)
    except:
        return {"campaign_name": "Campaign", "objective": "", "total_budget": 0,
                "start_date": str(date.today()), "end_date": str(date.today()),
                "audience": "", "channels": [], "assets": "", "market": "Sri Lanka"}

# ── Session defaults ──────────────────────────────────────────────────────────
for key,default in [("step",1),("brief",{}),("selected_client",None),
                    ("selected_brand",None),("combined_df",None),
                    ("generated_plan",None),("chat_messages",[]),
                    ("brief_complete",False),("brief_summary",{})]:
    if key not in st.session_state:
        st.session_state[key]=default

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='padding:8px 0 24px 0;'>
      <div style='font-family:Plus Jakarta Sans,sans-serif;font-size:1.25rem;font-weight:800;color:#1e3a5f;'>🎯 Orchestration</div>
      <div style='font-size:0.75rem;color:#8a93a8;margin-top:2px;'>Digital Planning Platform</div>
    </div>""", unsafe_allow_html=True)
    nav=st.radio("Navigation",["📊 New Campaign Plan","📁 Saved Plans","📈 Data Explorer"],label_visibility="collapsed")
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
# PAGE: NEW CAMPAIGN PLAN
# ══════════════════════════════════════════════════════════════════════════════
if nav=="📊 New Campaign Plan":
    sb=get_supabase()

    # Step indicator
    steps=["1 · Client & Brand","2 · Historical Data","3 · Plan Brief (Chat)","4 · Media Plan"]
    cols=st.columns(4)
    for i,(col,label) in enumerate(zip(cols,steps),1):
        active=st.session_state["step"]==i
        done=st.session_state["step"]>i
        bg="#1e3a5f" if active else ("#1a8a6e" if done else "#e8eaf0")
        fg="#ffffff" if (active or done) else "#8a93a8"
        icon="✓" if done else str(i)
        col.markdown(f"""<div style='background:{bg};color:{fg};border-radius:10px;padding:12px 16px;text-align:center;font-weight:600;font-size:0.82rem;'>{icon} · {label.split("·")[1].strip()}</div>""",unsafe_allow_html=True)

    st.markdown("<br>",unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 1: CLIENT & BRAND
    # ══════════════════════════════════════════════════════════════════════════
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
                    if r: st.success(f"✅ Client '{new_client}' created!"); st.rerun()
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
                        if r: st.success(f"✅ Brand '{new_brand}' created!"); st.rerun()
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
                st.session_state["brief_complete"]=False
                st.session_state["generated_plan"]=None
                st.rerun()
            st.markdown('</div>',unsafe_allow_html=True)

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 2: HISTORICAL DATA
    # ══════════════════════════════════════════════════════════════════════════
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
                    try:
                        dfs.append(pd.read_json(io.StringIO(f["data_json"])))
                    except: pass
            if dfs:
                st.session_state["combined_df"]=pd.concat(dfs,ignore_index=True)
        else:
            st.info(f"No historical data saved for **{brand_name}** yet. Upload below.")

        st.markdown('<div class="section-header">Upload New Files</div>',unsafe_allow_html=True)
        uploaded_files=st.file_uploader("Drop files here",type=["csv","xlsx","xls"],accept_multiple_files=True,help="Hold Ctrl/Cmd to select multiple")

        if uploaded_files:
            new_dfs=[]
            for f in uploaded_files:
                df=parse_uploaded_file(f)
                if df is not None: new_dfs.append((f.name,df))
            if new_dfs:
                if st.button(f"💾 Save {len(new_dfs)} file(s) to {brand_name}",use_container_width=True):
                    for fname,df in new_dfs:
                        save_brand_data(sb,brand_id,fname,df.to_json(),len(df))
                    st.success(f"✅ Saved!"); st.rerun()
                preview_df=pd.concat([d for _,d in new_dfs],ignore_index=True)
                with st.expander("Preview"):
                    st.dataframe(preview_df.head(20),use_container_width=True)
                existing=st.session_state.get("combined_df")
                st.session_state["combined_df"]=pd.concat(
                    ([existing]+[d for _,d in new_dfs]) if existing is not None else [d for _,d in new_dfs],
                    ignore_index=True)

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

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 3: PLANNING CHAT
    # ══════════════════════════════════════════════════════════════════════════
    elif st.session_state["step"]==3:
        client_name=st.session_state["selected_client"]["client_name"]
        brand_name=st.session_state["selected_brand"]["brand_name"]
        data_summary=summarise_dataframe(st.session_state["combined_df"])

        st.markdown(f'<div class="section-header">Planning Chat — {client_name} · {brand_name}</div>',unsafe_allow_html=True)
        st.caption("Chat with Orchy, your AI planning agent. Answer the questions to build your campaign brief.")

        # Initialise chat with agent greeting
        if not st.session_state["chat_messages"]:
            with st.spinner("Orchy is getting ready…"):
                first_msg=get_agent_response(
                    [{"role":"user","content":"Hello, I need to plan a new campaign."}],
                    client_name, brand_name, data_summary
                )
            st.session_state["chat_messages"]=[
                {"role":"user","content":"Hello, I need to plan a new campaign."},
                {"role":"agent","content":first_msg}
            ]
            st.rerun()

        # Render chat messages
        st.markdown('<div class="chat-container"><div class="chat-header"><p class="chat-header-title">🤖 Orchy — Planning Agent</p><p class="chat-header-sub">Ask me anything about your campaign</p></div><div class="chat-messages">',unsafe_allow_html=True)
        for msg in st.session_state["chat_messages"]:
            if msg["role"]!="system":
                render_message(msg["role"],msg["content"])
        st.markdown('</div></div>',unsafe_allow_html=True)

        # Check if brief is complete
        last_agent_msg=next((m["content"] for m in reversed(st.session_state["chat_messages"]) if m["role"]=="agent"),"")
        brief_complete="[BRIEF_COMPLETE]" in last_agent_msg

        if brief_complete:
            clean_last=last_agent_msg.replace("[BRIEF_COMPLETE]","").strip()
            if st.session_state["chat_messages"] and st.session_state["chat_messages"][-1]["content"]!=clean_last:
                st.session_state["chat_messages"][-1]["content"]=clean_last

            st.success("✅ Brief complete! Ready to generate your media plan.")
            st.markdown('<div class="green-btn">',unsafe_allow_html=True)
            if st.button("🚀 Generate Media Plan",use_container_width=True):
                with st.spinner("🤖 Analysing historical data and building your media plan…"):
                    plan_text=generate_media_plan(
                        st.session_state["chat_messages"],
                        client_name, brand_name, data_summary
                    )
                    brief_summary=extract_brief_from_conversation(
                        st.session_state["chat_messages"],
                        client_name, brand_name, data_summary
                    )
                    st.session_state["generated_plan"]=plan_text
                    st.session_state["brief_summary"]=brief_summary
                    st.session_state["step"]=4
                    st.rerun()
            st.markdown('</div>',unsafe_allow_html=True)
        else:
            # Input box
            with st.form("chat_form",clear_on_submit=True):
                col_inp,col_btn=st.columns([5,1])
                user_input=col_inp.text_input("Your message",placeholder="Type your answer here…",label_visibility="collapsed")
                submitted=col_btn.form_submit_button("Send →")

            if submitted and user_input.strip():
                st.session_state["chat_messages"].append({"role":"user","content":user_input.strip()})
                with st.spinner("Orchy is thinking…"):
                    agent_reply=get_agent_response(
                        st.session_state["chat_messages"],
                        client_name, brand_name, data_summary
                    )
                st.session_state["chat_messages"].append({"role":"agent","content":agent_reply})
                st.rerun()

        st.markdown("---")
        if st.button("← Back to Data"):
            st.session_state["step"]=2; st.rerun()

    # ══════════════════════════════════════════════════════════════════════════
    # STEP 4: MEDIA PLAN
    # ══════════════════════════════════════════════════════════════════════════
    elif st.session_state["step"]==4:
        client_name=st.session_state["selected_client"]["client_name"]
        brand_name=st.session_state["selected_brand"]["brand_name"]
        plan_text=st.session_state["generated_plan"]
        brief_summary=st.session_state.get("brief_summary",{})

        st.markdown(f'<div class="section-header">{client_name} · {brand_name} · {brief_summary.get("campaign_name","Media Plan")}</div>',unsafe_allow_html=True)

        budget=brief_summary.get("total_budget",0)
        try: budget=float(budget)
        except: budget=0

        c1,c2,c3,c4=st.columns(4)
        with c1: st.markdown(f'<div class="metric-card"><div class="metric-value">{format_lkr(budget)}</div><div class="metric-label">Total Budget</div></div>',unsafe_allow_html=True)
        with c2:
            try:
                s=datetime.strptime(brief_summary.get("start_date",""),"%Y-%m-%d")
                e=datetime.strptime(brief_summary.get("end_date",""),"%Y-%m-%d")
                days=(e-s).days
            except: days="—"
            st.markdown(f'<div class="metric-card"><div class="metric-value">{days} days</div><div class="metric-label">Duration</div></div>',unsafe_allow_html=True)
        with c3:
            chs=brief_summary.get("channels",[])
            st.markdown(f'<div class="metric-card"><div class="metric-value">{len(chs)}</div><div class="metric-label">Channels</div></div>',unsafe_allow_html=True)
        with c4:
            obj=brief_summary.get("objective","—")
            st.markdown(f'<div class="metric-card"><div class="metric-value">{obj.split()[0] if obj else "—"}</div><div class="metric-label">Objective</div></div>',unsafe_allow_html=True)

        plan_html=plan_text.replace("\n","<br>")
        st.markdown(f'<div class="plan-card">{plan_html}</div>',unsafe_allow_html=True)

        st.markdown("---")
        ca,cb,cc,cd=st.columns(4)
        with ca:
            st.download_button("📥 Download TXT",data=plan_text,
                file_name=f"{brief_summary.get('campaign_name','plan').replace(' ','_')}_plan.txt",
                mime="text/plain",use_container_width=True)
        with cb:
            excel_bytes=build_excel(brief_summary,plan_text,client_name,brand_name)
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
                    "total_budget":  budget,
                    "start_date":    brief_summary.get("start_date",""),
                    "end_date":      brief_summary.get("end_date",""),
                    "channels":      json.dumps(brief_summary.get("channels",[])),
                    "market":        brief_summary.get("market","Sri Lanka"),
                    "kpi_focus":     "",
                    "plan_text":     plan_text,
                    "created_at":    datetime.utcnow().isoformat()
                })
                st.success("Saved!")
        with cd:
            st.markdown('<div class="green-btn">',unsafe_allow_html=True)
            if st.button("➕ New Plan",use_container_width=True):
                st.session_state["step"]=1
                st.session_state["generated_plan"]=None
                st.session_state["chat_messages"]=[]
                st.session_state["brief_summary"]={}
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
                budget=plan.get("total_budget",0)
                try: budget=float(budget)
                except: budget=0
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