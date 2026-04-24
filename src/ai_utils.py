from __future__ import annotations

import json
import os
import re
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .metrics_utils import compact_metrics_for_ai


DEFAULT_OPENAI_MODEL = "gpt-4.1-mini"
DEFAULT_GEMINI_MODEL = "gemini-1.5-flash"
RULE_BASED_MODEL = "rules-v1"


@dataclass(frozen=True)
class NarrativeResult:
    text: str
    provider: str
    model: str
    used_api: bool
    warning: str = ""


def load_env_file(env_path: str | Path | None = None) -> None:
    path = Path(env_path) if env_path else Path(__file__).resolve().parents[1] / ".env"
    if not path.exists():
        return

    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and (key not in os.environ or os.environ[key] == ""):
            os.environ[key] = value


def normalize_provider(provider: str | None) -> str:
    value = (provider or "rule_based").strip().lower().replace("-", "_").replace(" ", "_")
    aliases = {
        "rule": "rule_based",
        "rules": "rule_based",
        "fallback": "rule_based",
        "rule_based_fallback": "rule_based",
        "open_ai": "openai",
        "google_gemini": "gemini",
    }
    return aliases.get(value, value if value in {"rule_based", "openai", "gemini"} else "rule_based")


def get_provider_status(env_path: str | Path | None = None) -> dict[str, dict[str, Any]]:
    load_env_file(env_path)
    return {
        "rule_based": {
            "label": "Rule-based fallback",
            "available": True,
            "model": RULE_BASED_MODEL,
        },
        "openai": {
            "label": "OpenAI",
            "available": bool(os.getenv("OPENAI_API_KEY")),
            "model": os.getenv("OPENAI_MODEL", DEFAULT_OPENAI_MODEL),
        },
        "gemini": {
            "label": "Gemini",
            "available": bool(os.getenv("GEMINI_API_KEY")),
            "model": os.getenv("GEMINI_MODEL", DEFAULT_GEMINI_MODEL),
        },
    }


def build_rule_based_narrative(metrics: dict[str, Any], period_label: str) -> str:
    facts = compact_metrics_for_ai(metrics)
    activities = list(facts.get("activity_breakdown", {}).keys())[:3]
    categories = list(facts.get("category_breakdown", {}).keys())[:3]
    quality = facts.get("quality_complexity", {})

    activity_text = ", ".join(activities) if activities else "no dominant activity type"
    category_text = ", ".join(categories) if categories else "no dominant reporting category"

    return (
        f"During {period_label}, the filtered data includes {facts['total_achievements']} achievements "
        f"across {facts['unique_studies']} studies or non-project workstreams, with contributions from "
        f"{facts['unique_contributors']} colleagues. The leading activity types are {activity_text}, "
        f"and the main reporting categories are {category_text}.\n\n"
        f"Quality and complexity indicators include {quality.get('high_impact_count', 0)} high-impact achievements, "
        f"{quality.get('quality_inspection_count', 0)} quality or inspection-related achievements, and "
        f"{facts['cf_total']} CF-related items. These metrics should be reviewed with the detailed table before "
        "final leadership distribution."
    )


def generate_narrative(
    *,
    metrics: dict[str, Any],
    period_label: str,
    provider: str = "rule_based",
    report_title: str = "CDO China Achievement Report",
    env_path: str | Path | None = None,
) -> NarrativeResult:
    load_env_file(env_path)
    requested_provider = normalize_provider(provider)
    fallback_text = build_rule_based_narrative(metrics, period_label)

    if requested_provider == "rule_based":
        return NarrativeResult(fallback_text, "rule_based", RULE_BASED_MODEL, used_api=False)

    status = get_provider_status(env_path).get(requested_provider, {})
    if not status.get("available"):
        warning = f"{status.get('label', requested_provider)} API key is not configured. Used rule-based fallback."
        return NarrativeResult(fallback_text, "rule_based", RULE_BASED_MODEL, used_api=False, warning=warning)

    try:
        system_prompt, user_prompt = _build_prompt(metrics, period_label, report_title)
        if requested_provider == "openai":
            model = str(status["model"])
            text = _call_openai(system_prompt, user_prompt, model)
        elif requested_provider == "gemini":
            model = str(status["model"])
            text = _call_gemini(system_prompt, user_prompt, model)
        else:
            return NarrativeResult(fallback_text, "rule_based", RULE_BASED_MODEL, used_api=False)

        text = text.strip()
        if not text:
            raise ValueError("AI provider returned an empty narrative.")

        if not _numbers_are_traceable(text, metrics, period_label, report_title):
            warning = "AI narrative contained numbers outside the calculated metrics. Used rule-based fallback."
            return NarrativeResult(fallback_text, "rule_based", RULE_BASED_MODEL, used_api=False, warning=warning)

        return NarrativeResult(text, requested_provider, model, used_api=True)
    except Exception as exc:
        warning = f"AI generation failed: {exc}. Used rule-based fallback."
        return NarrativeResult(fallback_text, "rule_based", RULE_BASED_MODEL, used_api=False, warning=warning)


def _build_prompt(metrics: dict[str, Any], period_label: str, report_title: str) -> tuple[str, str]:
    facts = compact_metrics_for_ai(metrics)
    facts_json = json.dumps(facts, ensure_ascii=False, indent=2, sort_keys=True, default=str)
    system_prompt = (
        "You write concise executive narratives for achievement reports. "
        "Use only the facts and numbers provided by the application metrics JSON. "
        "Do not invent counts, percentages, dates, rankings, study IDs, people, or source data. "
        "Do not change source data. Output editable plain text only, without markdown."
    )
    user_prompt = (
        f"Report title: {report_title}\n"
        f"Period: {period_label}\n"
        f"Application-calculated metrics JSON:\n{facts_json}\n\n"
        "Write 2 short paragraphs for an executive summary. "
        "If a metric is zero or missing, state it plainly or avoid emphasizing it. "
        "Every numeric value must appear exactly in the metrics JSON or period above."
    )
    return system_prompt, user_prompt


def _post_json(url: str, payload: dict[str, Any], headers: dict[str, str], timeout: int = 30) -> dict[str, Any]:
    data = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(url, data=data, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code}: {detail[:500]}") from exc


def _call_openai(system_prompt: str, user_prompt: str, model: str) -> str:
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY is not configured.")

    payload = {
        "model": model,
        "input": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.2,
        "max_output_tokens": 500,
    }
    data = _post_json(
        "https://api.openai.com/v1/responses",
        payload,
        {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
    )
    return _extract_openai_text(data)


def _extract_openai_text(data: dict[str, Any]) -> str:
    if data.get("output_text"):
        return str(data["output_text"])

    texts: list[str] = []
    for item in data.get("output", []):
        for content in item.get("content", []):
            text = content.get("text")
            if text:
                texts.append(str(text))
    return "\n".join(texts)


def _call_gemini(system_prompt: str, user_prompt: str, model: str) -> str:
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("GEMINI_API_KEY is not configured.")

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
    payload = {
        "contents": [
            {
                "role": "user",
                "parts": [{"text": f"{system_prompt}\n\n{user_prompt}"}],
            }
        ],
        "generationConfig": {
            "temperature": 0.2,
            "maxOutputTokens": 500,
        },
    }
    data = _post_json(url, payload, {"Content-Type": "application/json"})

    texts: list[str] = []
    for candidate in data.get("candidates", []):
        for part in candidate.get("content", {}).get("parts", []):
            text = part.get("text")
            if text:
                texts.append(str(text))
    return "\n".join(texts)


def _numbers_are_traceable(text: str, metrics: dict[str, Any], period_label: str, report_title: str) -> bool:
    allowed_source = json.dumps(
        {
            "metrics": compact_metrics_for_ai(metrics),
            "period_label": period_label,
            "report_title": report_title,
        },
        ensure_ascii=False,
        sort_keys=True,
        default=str,
    )
    allowed = {_normalize_number(token) for token in _extract_numbers(allowed_source)}
    return all(_normalize_number(token) in allowed for token in _extract_numbers(text))


def _extract_numbers(text: str) -> list[str]:
    return re.findall(r"(?<![A-Za-z])\d+(?:[,\d]*\d)?(?:\.\d+)?%?", text)


def _normalize_number(token: str) -> str:
    return token.replace(",", "").replace("%", "")
