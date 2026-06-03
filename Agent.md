# Agent.md — DeptFlow Achievement Reporter

Operating manual for AI coding agents (Copilot, Claude, Cursor, etc.) working on this
repository. Read this **before** making changes.

---

## 1. Project at a glance

- **Purpose:** Streamlit MVP that turns a SharePoint-exported "achievement" CSV into a
  filterable dashboard and a fixed 8-slide PowerPoint report. Includes an "AI Slide
  Studio" that proposes controlled JSON patches (text, fields, top-N, chart type, layout,
  style preset) which the user must explicitly apply before the preview/PPT changes.
- **Stack:** Python 3.11, Streamlit, pandas, plotly, python-pptx, openpyxl.
- **Entry point:** `app.py` (run via `streamlit run app.py`).
- **Deployment target:** Novartis OpenShift (container image built from this repo via Git
  import / S2I / BuildConfig). See `Dockerfile` and section 6 below.

## 2. Repository layout

```text
deptflow_achievement_demo/
├── app.py                       # Streamlit UI, orchestrates everything
├── requirements.txt             # Runtime Python deps
├── Dockerfile                   # OpenShift-compatible container build
├── .dockerignore
├── .streamlit/config.toml       # Headless / server settings for container
├── Agent.md                     # <— this file
├── README.md
├── .env.example                 # Template for local-only secrets
└── src/
    ├── data_utils.py            # CSV ingest, schema cleanup, filters
    ├── metrics_utils.py         # Dashboard metrics + report history trace
    ├── ai_utils.py              # AI provider abstraction (rule_based|openai|gemini)
    ├── slide_spec_utils.py      # SlideSpec model + patch apply + presets
    ├── slide_chat_utils.py      # AI chat → controlled JSON patch
    ├── html_preview_utils.py    # HTML/CSS slide preview (NOT rasterised)
    └── ppt_utils.py             # python-pptx renderer for the 8 fixed slides
```

Runtime-only / gitignored:

- `sample_data/` — local CSV exports.
- `outputs/` — generated `.pptx` and `report_history.csv`.
- `.env` — real API keys; only `.env.example` is committed.

## 3. Architectural rules (do not violate)

1. **Fixed 8-slide deck.** The PPT template is intentionally fixed:
   `Title → Executive Summary → Achievement Overview → Activity Type Breakdown →
   Project / Study Highlights → People Contribution View → Quality / Complexity
   Highlights → Appendix Detail Table`. Don't add or remove slides without an explicit
   product request; instead extend `SlideSpec` fields.
2. **Data integrity > AI.** All counts, filters, tables, and charts are computed by the
   app from the filtered DataFrame. AI is allowed to edit *narrative text only*. Never
   route a numeric value through an LLM.
3. **Patch-then-apply flow.** AI suggestions arrive as a JSON patch
   (`slide_chat_utils.propose_slide_patch`). Nothing changes the preview or the final PPT
   until the user clicks Apply (`slide_spec_utils.apply_slide_patch`). Preserve this gate.
4. **Optional AI provider.** `AI_PROVIDER=rule_based` must continue to work with **no**
   API keys. `openai` and `gemini` paths only activate when the matching key is present.
5. **HTML preview is not converted to images.** Preview and PPT are two renderers fed by
   the same `SlideSpec` + filtered data. Keep them in sync; don't try to screenshot the
   HTML for the deck.
6. **Traceability.** Every report build appends a row to `outputs/report_history.csv` via
   `metrics_utils.append_report_history` with a deterministic filter + spec signature.
   New features that change the output must extend the trace, not bypass it.

## 4. Coding conventions

- Python ≥ 3.11, `from __future__ import annotations` at the top of new modules.
- Type hints on all public functions; prefer `pathlib.Path` over raw strings for paths.
- Pure functions in `src/`; Streamlit/UI code stays in `app.py`.
- No global I/O at import time. `app.py` calls `load_env_file(ENV_PATH)` explicitly.
- Use `pandas` vectorised ops; avoid `DataFrame.iterrows` in hot paths.
- Don't introduce new heavyweight deps (e.g. matplotlib, torch, langchain) without a
  clear reason — the deck must build offline on a small OpenShift pod.
- Logging: prefer `st.info / st.warning / st.error` in UI; use `print` sparingly in
  utilities (container stdout is captured by OpenShift).

## 5. Local development

```powershell
cd C:\Flora\deptflow_achievement_demo
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
Copy-Item .env.example .env   # then edit if you need OpenAI/Gemini
streamlit run app.py
```

App opens on <http://localhost:8501>.

## 6. Container & OpenShift deployment

This repo is intended to be imported into Novartis OpenShift directly from Git
(BuildConfig + ImageStream, or "Import from Git" / S2I-style Docker strategy).

### Image contract

- **Base image:** `registry.access.redhat.com/ubi9/python-311` (Red Hat-maintained, FIPS-
  friendly, allowed in Novartis registries). Falls back to `python:3.11-slim` if the UBI
  registry is not reachable from your build cluster — adjust the `FROM` line accordingly.
- **Working dir:** `/opt/app-root/src` (UBI convention) or `/app` (slim fallback).
- **Listening port:** `8080` (non-privileged; matches the OpenShift default Service port
  convention). Streamlit is started with `--server.port=8080 --server.address=0.0.0.0
  --server.headless=true`.
- **User:** runs as a **non-root, arbitrary UID** in group `0` (root group). All
  application files are group-writable so OpenShift's random UID can read/write them.
  Do **not** add `USER 1001` with files chowned only to `1001:1001`.
- **Writable paths:** `/opt/app-root/src/outputs` is created and chmod `g+rwX` so report
  history and generated PPTX can be written. For persistence across pod restarts, mount a
  PVC at this path.
- **Healthcheck:** Streamlit exposes `/_stcore/health` returning `ok`. Use it for the
  OpenShift readiness/liveness probe.

### OpenShift Route / Service (suggested)

```yaml
# Service: targetPort 8080, port 8080
# Route:   TLS edge termination → Service:8080
# Probes:  HTTP GET /_stcore/health on 8080, initialDelay 10s, period 10s
```

### Secrets / config

Mount as environment variables (Secret + ConfigMap), **never** bake into the image:

| Variable                | Purpose                                          | Required |
| ----------------------- | ------------------------------------------------ | -------- |
| `AI_PROVIDER`           | `rule_based` (default) / `openai` / `gemini`     | no       |
| `OPENAI_API_KEY`        | only if `AI_PROVIDER=openai`                     | cond.    |
| `OPENAI_MODEL`          | default `gpt-4.1-mini`                           | no       |
| `GEMINI_API_KEY`        | only if `AI_PROVIDER=gemini`                     | cond.    |
| `GEMINI_MODEL`          | default `gemini-1.5-flash`                       | no       |
| `APP_ENV`               | `development` / `production`                     | no       |
| `DEFAULT_REPORT_TITLE`  | UI default                                       | no       |
| `SLIDE_STYLE_PRESET`    | matches keys in `slide_spec_utils.STYLE_PRESETS` | no       |

### Build via "Import from Git" (Docker strategy)

1. New Project → **+Add → Import from Git** → paste this repo's HTTPS URL.
2. Build strategy: **Dockerfile** (auto-detected from the repo root).
3. Target port: `8080`. Create a Route with TLS edge termination.
4. Add a Secret `deptflow-ai` with `OPENAI_API_KEY` / `GEMINI_API_KEY` if used, and
   reference it via `envFrom` on the Deployment.
5. (Optional) Attach a PVC of ≥ 1 Gi mounted at `/opt/app-root/src/outputs`.

## 7. Things an agent must NOT do

- ❌ Commit a real `.env`, API keys, or any file under `sample_data/` or `outputs/`.
- ❌ Hardcode a `USER 1001` (or any specific UID) in the Dockerfile — breaks OpenShift's
  random-UID security context.
- ❌ Bind to `localhost` / `127.0.0.1` inside the container; must be `0.0.0.0`.
- ❌ Add `EXPOSE 80` / `EXPOSE 443` or anything `< 1024` (privileged ports require root).
- ❌ Change the 8-slide template structure or let AI mutate numeric data.
- ❌ Add network calls at import time of any `src/*.py` module.

## 8. Useful test / smoke commands

```powershell
# Lint-ish sanity (no lint config yet — keep it minimal)
python -c "import app"                       # import-time errors?
python -c "from src import data_utils, ppt_utils, slide_spec_utils"

# Build container locally (Docker Desktop or Podman)
docker build -t deptflow-achievement:dev .
docker run --rm -p 8080:8080 deptflow-achievement:dev
# Open http://localhost:8080
```

## 9. Roadmap hooks (for context only)

- Add editable approval fields: `Reportable Flag`, `LM Confirmed`, `Achievement Title`,
  `Business Impact`, `Evidence Link`.
- Persist reviewed achievements to PostgreSQL (OpenShift-hosted) instead of CSV.
- Add AI services: comments → title, slide-category suggestion, business-impact wording,
  executive summary draft, duplicate detection. All routed through the
  `AI Suggestion → Human Review → Approved Narrative → PPT Output` gate.
