from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .ai_utils import _call_gemini, _call_openai, get_provider_status, load_env_file, normalize_provider
from .metrics_utils import compact_metrics_for_ai
from .slide_spec_utils import (
    DATA_SOURCE_FIELDS,
    PATCHABLE_SLIDE_FIELDS,
    STYLE_PRESETS,
    allowed_fields,
    validate_slide_patch,
)


@dataclass(frozen=True)
class SlidePatchResult:
    patch: dict[str, Any]
    provider: str
    model: str
    used_api: bool
    warning: str = ""


def propose_slide_patch(
    *,
    user_message: str,
    slide_spec: dict[str, Any],
    metrics: dict[str, Any],
    provider: str,
    env_path: str | Path | None = None,
) -> SlidePatchResult:
    load_env_file(env_path)
    requested_provider = normalize_provider(provider)
    status = get_provider_status(env_path).get(requested_provider, {})
    if requested_provider == "rule_based" or not status.get("available"):
        warning = "AI provider is not available. No patch was generated."
        return SlidePatchResult({}, "rule_based", "rules-v1", used_api=False, warning=warning)

    system_prompt, user_prompt = _build_patch_prompt(user_message, slide_spec, metrics)
    model = str(status["model"])
    try:
        if requested_provider == "openai":
            raw = _call_openai(system_prompt, user_prompt, model)
        elif requested_provider == "gemini":
            raw = _call_gemini(system_prompt, user_prompt, model)
        else:
            return SlidePatchResult({}, "rule_based", "rules-v1", used_api=False, warning="Unsupported provider.")

        patch = _parse_patch_json(raw)
        valid, errors = validate_slide_patch(patch, allowed_fields())
        if not valid:
            return SlidePatchResult({}, requested_provider, model, used_api=True, warning="Invalid patch: " + "; ".join(errors))
        return SlidePatchResult(patch, requested_provider, model, used_api=True)
    except Exception as exc:
        return SlidePatchResult({}, requested_provider, model, used_api=True, warning=f"AI slide patch failed: {exc}")


def _build_patch_prompt(user_message: str, slide_spec: dict[str, Any], metrics: dict[str, Any]) -> tuple[str, str]:
    safe_spec = json.dumps(slide_spec, ensure_ascii=False, indent=2, sort_keys=True, default=str)
    facts = json.dumps(compact_metrics_for_ai(metrics), ensure_ascii=False, indent=2, sort_keys=True, default=str)
    system_prompt = (
        "You are a controlled slide-spec editor. Return JSON only. "
        "Do not create, delete, or reorder slides. Do not invent or change numeric facts. "
        "You may only patch these slide fields: "
        f"{', '.join(sorted(PATCHABLE_SLIDE_FIELDS))}. "
        "Allowed style_preset values: "
        f"{', '.join(STYLE_PRESETS.keys())}. "
        "Allowed data_source fields map: "
        f"{json.dumps(DATA_SOURCE_FIELDS, ensure_ascii=False, sort_keys=True)}. "
        "The output schema is: {\"style_preset\": \"atlas|onyx|linen\", \"slides\": [{\"id\": 1, \"title\": \"...\"}]}. "
        "Omit unchanged fields. Use only fields listed in the data_source map."
    )
    user_prompt = (
        f"Current slide spec:\n{safe_spec}\n\n"
        f"Application-calculated metrics, for context only:\n{facts}\n\n"
        f"User request:\n{user_message}\n\n"
        "Return the smallest valid JSON patch. No markdown, no prose."
    )
    return system_prompt, user_prompt


def _parse_patch_json(raw: str) -> dict[str, Any]:
    text = raw.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?", "", text, flags=re.IGNORECASE).strip()
        text = re.sub(r"```$", "", text).strip()
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", text, flags=re.DOTALL)
        if not match:
            raise
        parsed = json.loads(match.group(0))
    if not isinstance(parsed, dict):
        raise ValueError("AI response must be a JSON object.")
    return parsed
