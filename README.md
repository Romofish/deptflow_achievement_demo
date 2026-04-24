# DeptFlow Achievement Reporter — Streamlit MVP

This demo turns a SharePoint-exported achievement CSV into a filterable dashboard and a fixed-template PowerPoint report.

## What it does

- Reads SharePoint achievement CSV files, including exports where the first row starts with `ListSchema=`.
- Cleans and standardizes fields.
- Supports filters by:
  - Delivery date range
  - Role
  - TA / Non-Project
  - Study ID
  - Person
  - Activity Type
  - Slide Category
  - Impact Level
- Shows dashboard metrics and charts.
- Builds a review table for human validation.
- Generates an 8-slide fixed-template PPTX report from filtered data only.
- Includes an editable AI Narrative Assistant with rule-based fallback.
- Records report generation traceability in `outputs/report_history.csv`.

## Folder structure

```text
/deptflow_achievement_demo
  app.py
  requirements.txt
  README.md
  /src
    data_utils.py
    metrics_utils.py
    ai_utils.py
    ppt_utils.py
  /sample_data
    local CSV exports, ignored by Git
  /outputs
```

`sample_data/`, `outputs/`, virtual environments, and real `.env` files are ignored by Git. This keeps exported SharePoint data, generated PPTX files, and API keys out of version control.

## How to run

```bash
cd deptflow_achievement_demo
python -m venv .venv

# macOS / Linux
source .venv/bin/activate

# Windows PowerShell
# .venv\Scripts\Activate.ps1

python -m pip install --upgrade pip
pip install -r requirements.txt
streamlit run app.py
```

If the bundled sample CSV is not present, upload a SharePoint achievement CSV in the sidebar. PPTX export works from the uploaded and filtered data.

For AI provider usage, copy `.env.example` to `.env` and fill local-only settings such as `OPENAI_API_KEY` or `GEMINI_API_KEY`. The real `.env` file is intentionally not committed.

AI narrative generation is optional:

- `AI_PROVIDER=rule_based` works without API keys.
- `AI_PROVIDER=openai` uses OpenAI only when `OPENAI_API_KEY` is present.
- `AI_PROVIDER=gemini` uses Gemini only when `GEMINI_API_KEY` is present.

AI output is editable narrative text only. Counts, filters, and PPT tables/charts are calculated by the application from the filtered source data.

## Recommended next step

For the next version, add editable approval fields:

- `Reportable Flag`
- `LM Confirmed`
- `Achievement Title`
- `Business Impact`
- `Evidence Link`

Then save the reviewed data into a database such as PostgreSQL / Supabase / OpenShift-hosted PostgreSQL.

## AI extension idea

The MVP intentionally does not require an AI API key. Later, add an AI service for:

- comments-to-achievement-title
- slide category suggestion
- business impact wording
- executive summary drafting
- duplicate achievement detection

Recommended control flow:

```text
AI Suggestion → Human Review → Approved Narrative → PPT Output
```
