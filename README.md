# 🎯 Orchestration-Digital

> AI-powered digital campaign planning platform

## What it does
- Upload past campaign data (CSV/Excel from Meta, Google, TikTok, etc.)
- Fill in your campaign brief
- Get an AI-generated media plan with budget splits, channel tactics, KPI targets
- Save plans to your library and export as Excel or TXT

## Tech Stack
- **Frontend:** Streamlit
- **AI:** Anthropic Claude (claude-sonnet-4-6)
- **Database:** Supabase (Postgres)
- **Hosting:** Streamlit Community Cloud

## Setup

### 1. Supabase — Create the database table
Run this SQL in your Supabase SQL editor:

```sql
CREATE TABLE campaign_plans (
  id uuid DEFAULT gen_random_uuid() PRIMARY KEY,
  campaign_name text,
  objective text,
  total_budget numeric,
  start_date text,
  end_date text,
  channels text,
  market text,
  kpi_focus text,
  plan_text text,
  created_at timestamp with time zone DEFAULT now()
);
```

### 2. Streamlit Cloud — Add secrets
In your Streamlit Cloud app settings → Secrets, paste:

```toml
ANTHROPIC_API_KEY = "your-key"
SUPABASE_URL = "https://xxxx.supabase.co"
SUPABASE_KEY = "your-anon-key"
```

### 3. Deploy
Connect this GitHub repo to Streamlit Cloud and deploy `app.py`.

## Phase Roadmap
- **Phase 1 (now):** Upload CSV/Excel → AI plans
- **Phase 2:** Direct API integrations (Meta, Google, TikTok)
- **Phase 3:** Execution, monitoring & reporting dashboard
