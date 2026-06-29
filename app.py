import streamlit as st
import pandas as pd
import json
import anthropic
from datetime import datetime, date
import io
from supabase import create_client, Client
import os

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Orchestration-Digital",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Space+Grotesk:wght@400;500;700&display=swap');

  html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
  }

  /* Background */
  .stApp {
    background: #0d0f1a;
    color: #e8eaf0;
  }

  /* Sidebar */
  [data-testid="stSidebar"] {
    background: #12152b !important;
    border-right: 1px solid #1e2340;
  }

  /* Hero header */
  .hero-header {
    background: linear-gradient(135deg, #1a1f3e 0%, #0d1526 100%);
    border: 1px solid #2a3060;
    border-radius: 16px;
    padding: 36px 40px;
    margin-bottom: 32px;
    position: relative;
    overflow: hidden;
  }
  .hero-header::before {
    content: '';
    position: absolute;
    top: -60px; right: -60px;
    width: 220px; height: 220px;
    background: radial-gradient(circle, rgba(99,102,241,0.15) 0%, transparent 70%);
    border-radius: 50%;
  }
  .hero-title {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 2rem;
    font-weight: 700;
    color: #ffffff;
    letter-spacing: -0.5px;
    margin: 0 0 6px 0;
  }
  .hero-sub {
    font-size: 0.95rem;
    color: #7b82a8;
    margin: 0;
  }
  .hero-badge {
    display: inline-block;
    background: rgba(99,102,241,0.15);
    color: #818cf8;
    border: 1px solid rgba(99,102,241,0.3);
    border-radius: 20px;
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    padding: 4px 12px;
    margin-bottom: 14px;
  }

  /* Metric cards */
  .metric-card {
    background: #12152b;
    border: 1px solid #1e2340;
    border-radius: 12px;
    padding: 20px 24px;
    text-align: center;
  }
  .metric-value {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 1.8rem;
    font-weight: 700;
    color: #818cf8;
  }
  .metric-label {
    font-size: 0.8rem;
    color: #5c6285;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin-top: 4px;
  }

  /* Section headers */
  .section-header {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 1.1rem;
    font-weight: 600;
    color: #c7cae0;
    border-left: 3px solid #6366f1;
    padding-left: 12px;
    margin: 28px 0 16px 0;
  }

  /* Plan output card */
  .plan-card {
    background: #12152b;
    border: 1px solid #1e2340;
    border-radius: 14px;
    padding: 28px 32px;
    margin-top: 20px;
    line-height: 1.75;
    color: #c7cae0;
  }

  /* Streamlit widget overrides */
  .stButton > button {
    background: linear-gradient(135deg, #6366f1, #4f46e5) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    padding: 10px 24px !important;
    font-size: 0.9rem !important;
    transition: opacity 0.2s !important;
  }
  .stButton > button:hover { opacity: 0.85 !important; }

  .stTextInput > div > div > input,
  .stNumberInput > div > div > input,
  .stSelectbox > div > div,
  .stDateInput > div > div > input {
    background: #1a1f3e !important;
    border: 1px solid #2a3060 !important;
    color: #e8eaf0 !important;
    border-radius: 8px !important;
  }

  [data-testid="stFileUploader"] {
    background: #12152b !important;
    border: 2px dashed #2a3060 !important;
    border-radius: 12px !important;
  }

  .stTabs [data-baseweb="tab-list"] {
    background: #12152b;
    border-radius: 10px;
    padding: 4px;
    gap: 4px;
  }
  .stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: #5c6285 !important;
    border-radius: 8px !important;
    font-weight: 500 !important;
  }
  .stTabs [aria-selected="true"] {
    background: #6366f1 !important;
    color: white !important;
  }

  div[data-testid="stExpander"] {
    background: #12152b !important;
    border: 1px solid #1e2340 !important;
    border-radius: 10px !important;
  }

  .stAlert { border-radius: 10px !important; }

  /* Hide Streamlit branding */
  #MainMenu, footer, header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Clients ───────────────────────────────────────────────────────────────────
@st.cache_resource
def get_supabase() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

def get_anthropic():
    return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])

# ── Helpers ───────────────────────────────────────────────────────────────────
def parse_uploaded_file(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    elif name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file)
    else:
        st.error("Please upload a CSV or Excel file.")
        return None

def summarise_dataframe(df: pd.DataFrame) -> str:
    summary_lines = [
        f"Rows: {len(df)}, Columns: {len(df.columns)}",
        f"Columns: {', '.join(df.columns.tolist())}",
        "",
        "Sample data (first 10 rows):",
        df.head(10).to_string(index=False),
        "",
        "Numeric column statistics:",
        df.describe().to_string()
    ]
    return "\n".join(summary_lines)

def generate_campaign_plan(
    client: anthropic.Anthropic,
    data_summary: str,
    campaign_name: str,
    objective: str,
    total_budget: float,
    start_date: date,
    end_date: date,
    target_audience: str,
    channels: list,
    market: str,
    kpi_focus: str
) -> str:
    duration_days = (end_date - start_date).days
    channel_str = ", ".join(channels) if channels else "All channels"

    prompt = f"""You are a senior digital media planner with 15+ years of experience planning campaigns across Search, Social, Programmatic, and Video channels.

Based on the historical campaign performance data below, create a detailed digital media plan for the following brief:

--- CAMPAIGN BRIEF ---
Campaign Name: {campaign_name}
Objective: {objective}
Total Budget: ${total_budget:,.2f}
Flight Dates: {start_date.strftime('%d %b %Y')} – {end_date.strftime('%d %b %Y')} ({duration_days} days)
Target Audience: {target_audience}
Preferred Channels: {channel_str}
Market: {market}
Primary KPI: {kpi_focus}

--- HISTORICAL PERFORMANCE DATA ---
{data_summary}

--- YOUR TASK ---
Using the historical data to inform benchmarks and budget efficiency, produce a complete media plan with these sections:

1. EXECUTIVE SUMMARY
   Brief strategic rationale (3–4 sentences) for the recommended approach.

2. CHANNEL MIX & BUDGET ALLOCATION
   For each recommended channel, provide:
   - Budget allocation ($ amount and % of total)
   - Rationale based on historical data
   - Expected reach or impressions
   - Flight dates within the campaign window

3. CHANNEL-BY-CHANNEL TACTICS
   For each channel:
   - Ad formats recommended
   - Targeting approach
   - Bidding strategy
   - Creative recommendations

4. KPI TARGETS & BENCHMARKS
   Based on historical data, set realistic targets for:
   - Primary KPI: {kpi_focus}
   - Secondary KPIs (CTR, CPM, CPC, ROAS as relevant)
   - Weekly pacing milestones

5. OPTIMISATION ROADMAP
   Week-by-week optimisation actions to maximise performance.

6. RISK FLAGS
   Any risks identified from the historical data and mitigation strategies.

Format your response clearly with headers and bullet points. Be specific with numbers — use the historical data to justify every budget and KPI recommendation."""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )
    return message.content[0].text

def save_plan_to_supabase(supabase: Client, plan_data: dict):
    try:
        result = supabase.table("campaign_plans").insert(plan_data).execute()
        return True
    except Exception as e:
        st.warning(f"Could not save to database: {e}")
        return False

def load_saved_plans(supabase: Client):
    try:
        result = supabase.table("campaign_plans").select("*").order("created_at", desc=True).limit(20).execute()
        return result.data
    except Exception as e:
        return []

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='padding: 8px 0 24px 0;'>
        <div style='font-family: Space Grotesk, sans-serif; font-size: 1.3rem; font-weight: 700; color: #fff;'>
            🎯 Orchestration
        </div>
        <div style='font-size: 0.75rem; color: #5c6285; margin-top: 2px;'>Digital Planning Platform</div>
    </div>
    """, unsafe_allow_html=True)

    nav = st.radio(
        "Navigation",
        ["📊 New Campaign Plan", "📁 Saved Plans", "📈 Data Explorer"],
        label_visibility="collapsed"
    )

    st.markdown("---")
    st.markdown("<div style='font-size:0.75rem; color:#5c6285;'>POC Version 1.0<br>Phase 1 — Upload & Plan</div>", unsafe_allow_html=True)

# ── Hero ──────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero-header">
  <div class="hero-badge">Digital Planning Platform</div>
  <div class="hero-title">Orchestration-Digital</div>
  <div class="hero-sub">Upload past campaign data · Generate AI-powered media plans · Export &amp; execute</div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: NEW CAMPAIGN PLAN
# ══════════════════════════════════════════════════════════════════════════════
if nav == "📊 New Campaign Plan":

    tab1, tab2, tab3 = st.tabs(["  1 · Upload Data  ", "  2 · Campaign Brief  ", "  3 · Your Plan  "])

    # ── Tab 1: Upload ─────────────────────────────────────────────────────────
    with tab1:
        st.markdown('<div class="section-header">Upload Historical Campaign Data</div>', unsafe_allow_html=True)
        st.caption("Upload CSV or Excel exports from Meta, Google, TikTok, or any ad platform. The AI will learn from your past results.")

        uploaded_file = st.file_uploader(
            "Drop your campaign file here",
            type=["csv", "xlsx", "xls"],
            help="Supported: CSV, XLSX, XLS"
        )

        if uploaded_file:
            df = parse_uploaded_file(uploaded_file)
            if df is not None:
                st.session_state["uploaded_df"] = df
                st.session_state["file_name"] = uploaded_file.name

                st.success(f"✅ **{uploaded_file.name}** loaded — {len(df):,} rows × {len(df.columns)} columns")

                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown(f'<div class="metric-card"><div class="metric-value">{len(df):,}</div><div class="metric-label">Data Rows</div></div>', unsafe_allow_html=True)
                with col2:
                    st.markdown(f'<div class="metric-card"><div class="metric-value">{len(df.columns)}</div><div class="metric-label">Columns</div></div>', unsafe_allow_html=True)
                with col3:
                    numeric_cols = df.select_dtypes(include='number').columns
                    st.markdown(f'<div class="metric-card"><div class="metric-value">{len(numeric_cols)}</div><div class="metric-label">Numeric Fields</div></div>', unsafe_allow_html=True)

                with st.expander("Preview data"):
                    st.dataframe(df.head(20), use_container_width=True)

                with st.expander("Column summary"):
                    st.dataframe(df.describe(include='all').T, use_container_width=True)

                st.info("✅ Data ready. Go to **2 · Campaign Brief** to build your plan.")

        else:
            st.markdown("""
            <div style='text-align:center; padding: 40px; color: #5c6285;'>
                <div style='font-size: 2.5rem; margin-bottom: 12px;'>📂</div>
                <div style='font-weight: 500; margin-bottom: 6px;'>No file uploaded yet</div>
                <div style='font-size: 0.85rem;'>Export your data from Meta Ads Manager, Google Ads, or TikTok Ads and upload above</div>
            </div>
            """, unsafe_allow_html=True)

    # ── Tab 2: Brief ──────────────────────────────────────────────────────────
    with tab2:
        if "uploaded_df" not in st.session_state:
            st.warning("⬅️ Please upload your campaign data first in the **Upload Data** tab.")
        else:
            st.markdown('<div class="section-header">Campaign Brief</div>', unsafe_allow_html=True)
            st.caption("Fill in your campaign details and we'll generate a data-driven media plan.")

            col_a, col_b = st.columns(2)

            with col_a:
                campaign_name = st.text_input("Campaign Name", placeholder="e.g. Q3 2025 Brand Awareness")
                objective = st.selectbox("Campaign Objective", [
                    "Brand Awareness", "Reach & Frequency", "Website Traffic",
                    "Lead Generation", "App Installs", "E-commerce / Conversions",
                    "Video Views", "Engagement"
                ])
                total_budget = st.number_input("Total Budget (USD $)", min_value=100.0, value=10000.0, step=500.0)
                market = st.text_input("Market / Region", placeholder="e.g. Sri Lanka, Southeast Asia, Global")

            with col_b:
                target_audience = st.text_area("Target Audience", placeholder="e.g. 25–45, urban professionals, interested in finance & tech", height=100)
                kpi_focus = st.selectbox("Primary KPI", [
                    "ROAS (Return on Ad Spend)", "CPA (Cost per Acquisition)",
                    "CPL (Cost per Lead)", "CTR (Click-through Rate)",
                    "CPM (Cost per 1000 Impressions)", "CPC (Cost per Click)",
                    "App Installs", "Video Completion Rate"
                ])
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    start_date = st.date_input("Start Date", value=date.today())
                with col_d2:
                    end_date = st.date_input("End Date")

            st.markdown('<div class="section-header">Channel Selection</div>', unsafe_allow_html=True)
            col_c1, col_c2, col_c3, col_c4 = st.columns(4)
            with col_c1:
                ch_meta = st.checkbox("Meta (Facebook/Instagram)", value=True)
            with col_c2:
                ch_google = st.checkbox("Google Ads", value=True)
            with col_c3:
                ch_tiktok = st.checkbox("TikTok Ads", value=True)
            with col_c4:
                ch_programmatic = st.checkbox("Programmatic Display", value=False)

            col_c5, col_c6 = st.columns(4)[:2]
            with col_c5:
                ch_youtube = st.checkbox("YouTube", value=False)
            with col_c6:
                ch_linkedin = st.checkbox("LinkedIn", value=False)

            channels_selected = []
            if ch_meta: channels_selected.append("Meta (Facebook & Instagram)")
            if ch_google: channels_selected.append("Google Search & Display")
            if ch_tiktok: channels_selected.append("TikTok Ads")
            if ch_programmatic: channels_selected.append("Programmatic Display")
            if ch_youtube: channels_selected.append("YouTube")
            if ch_linkedin: channels_selected.append("LinkedIn")

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
                        "campaign_name": campaign_name,
                        "objective": objective,
                        "total_budget": total_budget,
                        "start_date": start_date,
                        "end_date": end_date,
                        "target_audience": target_audience,
                        "channels": channels_selected,
                        "market": market,
                        "kpi_focus": kpi_focus
                    }

                    with st.spinner("🤖 Analysing your historical data and building the plan…"):
                        try:
                            client = get_anthropic()
                            data_summary = summarise_dataframe(st.session_state["uploaded_df"])
                            plan_text = generate_campaign_plan(
                                client, data_summary, campaign_name, objective,
                                total_budget, start_date, end_date,
                                target_audience, channels_selected, market, kpi_focus
                            )
                            st.session_state["generated_plan"] = plan_text
                            st.success("✅ Plan generated! Go to **3 · Your Plan** to review it.")
                        except Exception as e:
                            st.error(f"Error generating plan: {e}")

    # ── Tab 3: Plan ───────────────────────────────────────────────────────────
    with tab3:
        if "generated_plan" not in st.session_state:
            st.info("👆 Complete the brief in **2 · Campaign Brief** and hit Generate to see your plan here.")
        else:
            brief = st.session_state.get("brief", {})

            st.markdown(f"""
            <div style='display:flex; gap:12px; flex-wrap:wrap; margin-bottom:24px;'>
                <div class="metric-card" style='flex:1; min-width:140px;'>
                    <div class="metric-value">${brief.get('total_budget',0):,.0f}</div>
                    <div class="metric-label">Total Budget</div>
                </div>
                <div class="metric-card" style='flex:1; min-width:140px;'>
                    <div class="metric-value">{(brief.get('end_date', date.today()) - brief.get('start_date', date.today())).days}</div>
                    <div class="metric-label">Campaign Days</div>
                </div>
                <div class="metric-card" style='flex:1; min-width:140px;'>
                    <div class="metric-value">{len(brief.get('channels', []))}</div>
                    <div class="metric-label">Channels</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            st.markdown(f'<div class="plan-card">{st.session_state["generated_plan"].replace(chr(10), "<br>")}</div>', unsafe_allow_html=True)

            st.markdown("---")
            col_e1, col_e2, col_e3 = st.columns(3)

            with col_e1:
                plan_text = st.session_state["generated_plan"]
                st.download_button(
                    "📥 Download as TXT",
                    data=plan_text,
                    file_name=f"{brief.get('campaign_name','plan').replace(' ','_')}_plan.txt",
                    mime="text/plain",
                    use_container_width=True
                )

            with col_e2:
                plan_df = pd.DataFrame([{
                    "Campaign": brief.get("campaign_name"),
                    "Objective": brief.get("objective"),
                    "Budget": brief.get("total_budget"),
                    "Start": brief.get("start_date"),
                    "End": brief.get("end_date"),
                    "Channels": ", ".join(brief.get("channels", [])),
                    "Market": brief.get("market"),
                    "KPI": brief.get("kpi_focus"),
                    "Plan": plan_text
                }])
                excel_buf = io.BytesIO()
                plan_df.to_excel(excel_buf, index=False)
                st.download_button(
                    "📊 Download as Excel",
                    data=excel_buf.getvalue(),
                    file_name=f"{brief.get('campaign_name','plan').replace(' ','_')}_plan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            with col_e3:
                if st.button("💾 Save to Library", use_container_width=True):
                    supabase = get_supabase()
                    save_plan_to_supabase(supabase, {
                        "campaign_name": brief.get("campaign_name"),
                        "objective": brief.get("objective"),
                        "total_budget": brief.get("total_budget"),
                        "start_date": str(brief.get("start_date")),
                        "end_date": str(brief.get("end_date")),
                        "channels": json.dumps(brief.get("channels", [])),
                        "market": brief.get("market"),
                        "kpi_focus": brief.get("kpi_focus"),
                        "plan_text": plan_text,
                        "created_at": datetime.utcnow().isoformat()
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
                with st.expander(f"📋 {plan.get('campaign_name', 'Unnamed')} — {plan.get('market', '')} — ${plan.get('total_budget', 0):,.0f}"):
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Objective", plan.get("objective", "—"))
                    col2.metric("Budget", f"${plan.get('total_budget', 0):,.0f}")
                    col3.metric("KPI", plan.get("kpi_focus", "—"))
                    st.markdown(f"**Channels:** {plan.get('channels', '')}")
                    st.markdown(f"**Dates:** {plan.get('start_date')} → {plan.get('end_date')}")
                    st.markdown("---")
                    st.markdown(plan.get("plan_text", ""))
    except Exception as e:
        st.error(f"Could not load plans: {e}. Please check your Supabase connection.")

# ══════════════════════════════════════════════════════════════════════════════
# PAGE: DATA EXPLORER
# ══════════════════════════════════════════════════════════════════════════════
elif nav == "📈 Data Explorer":
    st.markdown('<div class="section-header">Data Explorer</div>', unsafe_allow_html=True)

    if "uploaded_df" not in st.session_state:
        st.info("Upload campaign data in the **New Campaign Plan** section first.")
    else:
        df = st.session_state["uploaded_df"]
        st.caption(f"Exploring: **{st.session_state.get('file_name', 'uploaded file')}** — {len(df):,} rows")

        st.dataframe(df, use_container_width=True)

        numeric_cols = df.select_dtypes(include='number').columns.tolist()
        if len(numeric_cols) >= 2:
            st.markdown('<div class="section-header">Quick Chart</div>', unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                x_col = st.selectbox("X axis", df.columns.tolist())
            with col2:
                y_col = st.selectbox("Y axis", numeric_cols)

            chart_type = st.radio("Chart type", ["Line", "Bar", "Area"], horizontal=True)

            chart_df = df[[x_col, y_col]].dropna().set_index(x_col)
            if chart_type == "Line":
                st.line_chart(chart_df)
            elif chart_type == "Bar":
                st.bar_chart(chart_df)
            else:
                st.area_chart(chart_df)
