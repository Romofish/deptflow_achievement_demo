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
- Includes an AI Slide Studio with HTML/CSS slide preview and controlled AI patches.
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
    slide_spec_utils.py
    slide_chat_utils.py
    html_preview_utils.py
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

## HTML preview and PPT output

The app previews slides as HTML/CSS inside Streamlit. This preview is not converted into images. When you build the report, the app uses the same `SlideSpec` plus filtered data to generate a real editable `.pptx` file with PowerPoint text, tables, charts, and shapes.

The fixed deck keeps 8 slides:

1. Title Page
2. Executive Summary
3. Achievement Overview Dashboard
4. Activity Type Breakdown
5. Project / Study Highlights
6. People Contribution View
7. Quality / Complexity Highlights
8. Appendix Detail Table

AI Slide Studio can propose controlled JSON patches for slide text, fields, top-N limits, chart type, layout variant, and style preset. The user must apply the patch before it changes the preview or final PPT.

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
