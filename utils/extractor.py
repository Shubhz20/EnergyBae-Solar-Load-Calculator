"""
Gemini 1.5 Flash extractor for MSEDCL electricity bills.

Returns a JSON object with:
    consumer_name, consumer_number, billing_unit,
    connected_load_kw, tariff_category, avg_monthly_consumption
"""

import hashlib
import json
import mimetypes
import os
import re
import threading
from pathlib import Path
from typing import Any, Dict, Optional

import google.generativeai as genai
from google.api_core import exceptions as gax_exceptions


# In-process cache of {file_sha256: extraction_dict} so re-running the same
# bill is free (matters when free-tier quota is tight).
_CACHE: Dict[str, Dict[str, Any]] = {}
_CACHE_LOCK = threading.Lock()

PROMPT = """
You are an expert document parser specialised in MSEDCL (Maharashtra State
Electricity Distribution Co. Ltd.) electricity bills.

Extract the following fields from the bill and return STRICT JSON only.
Do not include markdown fences, comments, or explanations.

Required JSON schema:
{
  "consumer_name": string,            // billed customer's full name
  "consumer_number": string,          // 12-digit MSEDCL consumer number
  "billing_unit": string,             // BU code or BU name printed on the bill
  "connected_load_kw": number,        // sanctioned/connected load in kW.
                                      // If the bill prints HP, convert: 1 HP = 0.7457 kW.
                                      // If it prints kVA, keep the numeric value but note
                                      // the unit by still returning kW (kVA ~ kW for PF=1).
  "tariff_category": string,          // e.g. "LT-I Residential", "LT-II Commercial",
                                      // "LT-V Industrial", "HT-I", etc.
  "avg_monthly_consumption": number   // average monthly units (kWh) over the
                                      // last 6 months. If a 6-month consumption
                                      // history table is present, compute the
                                      // arithmetic mean. Otherwise use the
                                      // current month's units.
}

Rules:
- Return numbers as numbers, not strings.
- If a field is not found, use null.
- Never invent values. Prefer null over guessing.
- Output JSON only.
""".strip()


def _load_file_part(file_path: str) -> Dict[str, Any]:
    """Read a local file into a Gemini inline part."""
    path = Path(file_path)
    mime, _ = mimetypes.guess_type(path.name)
    if mime is None:
        # Default to PDF; Gemini also accepts image/png, image/jpeg
        mime = "application/pdf"
    return {"mime_type": mime, "data": path.read_bytes()}


def _coerce_json(text: str) -> Dict[str, Any]:
    """Strip code fences / stray text and parse JSON."""
    cleaned = text.strip()
    # Drop ```json ... ``` fences if present
    cleaned = re.sub(r"^```(?:json)?", "", cleaned).strip()
    cleaned = re.sub(r"```$", "", cleaned).strip()

    # Fall back: grab the first {...} block
    if not cleaned.startswith("{"):
        match = re.search(r"\{.*\}", cleaned, flags=re.DOTALL)
        if match:
            cleaned = match.group(0)

    return json.loads(cleaned)


def _normalise(data: Dict[str, Any]) -> Dict[str, Any]:
    """Coerce numeric fields and ensure all keys exist."""
    keys = [
        "consumer_name",
        "consumer_number",
        "billing_unit",
        "connected_load_kw",
        "tariff_category",
        "avg_monthly_consumption",
    ]
    out = {k: data.get(k) for k in keys}

    for num_key in ("connected_load_kw", "avg_monthly_consumption"):
        v = out.get(num_key)
        if isinstance(v, str):
            m = re.search(r"-?\d+(?:\.\d+)?", v.replace(",", ""))
            out[num_key] = float(m.group(0)) if m else None
    return out


def extract_bill_data(file_path: str, api_key: str) -> Dict[str, Any]:
    """
    Send the bill (PDF or image) to Gemini Flash and return the parsed
    field dictionary. Walks live model candidates and retries on 404 /
    deprecation errors.
    """
    if not api_key:
        raise ValueError("Gemini API key is required.")

    api_key = api_key.strip()
    if not api_key.startswith("AIza"):
        raise ValueError(
            "That doesn't look like a Gemini API key. Keys from "
            "https://aistudio.google.com/app/apikey start with 'AIza'."
        )

    genai.configure(api_key=api_key)

    # Cache lookup - identical bill -> identical extraction, no API call.
    file_bytes = Path(file_path).read_bytes()
    cache_key = hashlib.sha256(file_bytes + api_key.encode()).hexdigest()
    with _CACHE_LOCK:
        if cache_key in _CACHE:
            return _CACHE[cache_key]

    file_part = _load_file_part(file_path)

    candidates = _candidate_models()
    last_err: Optional[Exception] = None
    quota_errors: list[str] = []

    for name in candidates:
        try:
            model = genai.GenerativeModel(name)
            response = model.generate_content(
                [PROMPT, file_part],
                generation_config={
                    "temperature": 0.0,
                    "response_mime_type": "application/json",
                },
            )
            raw = response.text or ""
            data = _coerce_json(raw)
            data = _normalise(data)
            with _CACHE_LOCK:
                _CACHE[cache_key] = data
            return data

        except gax_exceptions.NotFound as e:
            # Model name doesn't exist or isn't callable.
            last_err = e
            continue
        except gax_exceptions.FailedPrecondition as e:
            # Model deprecated for this key/region.
            last_err = e
            continue
        except gax_exceptions.ResourceExhausted as e:
            # 429 - quota exhausted on this model. Each model has its own
            # quota bucket, so try the next candidate.
            last_err = e
            quota_errors.append(name)
            continue
        except Exception as e:  # noqa: BLE001
            raise

    # If we got here every Gemini call failed. Try the offline parser for
    # PDFs as a last resort - zero API, zero quota.
    if Path(file_path).suffix.lower() == ".pdf":
        try:
            from utils.regex_extractor import extract_from_pdf
            data = extract_from_pdf(file_path)
            data["_method"] = "offline_regex"
            with _CACHE_LOCK:
                _CACHE[cache_key] = data
            return data
        except Exception as fallback_err:  # noqa: BLE001
            last_err = f"Gemini failed and offline parser also failed: {fallback_err}"

    if quota_errors and len(quota_errors) == len(candidates):
        raise RuntimeError(
            "All Gemini Flash models hit free-tier quota AND the offline "
            "parser couldn't read this file.\n\n"
            "Fixes:\n"
            "  - For PDF bills: the offline parser should normally work. "
            "Check that the PDF has selectable text (not a scan).\n"
            "  - For images/scans: you need a working Gemini key. Create a "
            "fresh one at https://aistudio.google.com/app/apikey using "
            "'Create API key in new project'.\n"
            f"Models tried: {quota_errors}"
        )

    raise RuntimeError(
        f"No working Gemini Flash model for this API key. "
        f"Tried: {candidates}. Last error: {last_err}"
    )


def _candidate_models() -> list[str]:
    """
    Build a prioritised list of model names to try.

    Priority:
        1. Explicit override via GEMINI_MODEL env var (admin escape hatch).
        2. Whatever live Flash models list_models() reports as supporting
           generateContent, ordered: 2.0 Flash -> 2.5 Flash -> 1.5 Flash
           (1.5 is deprecated as of 2026), preferring "-latest" aliases.
        3. Hard-coded fallback chain in case list_models() fails.
    """
    override = os.environ.get("GEMINI_MODEL", "").strip()
    if override:
        return [override]

    try:
        live = list(genai.list_models())
        flash = [
            m for m in live
            if "flash" in getattr(m, "name", "").lower()
            and "generateContent" in (getattr(m, "supported_generation_methods", []) or [])
        ]

        def rank(m) -> tuple:
            n = m.name.lower()
            gen = (
                0 if "2.0-flash" in n
                else 1 if "2.5-flash" in n
                else 2 if "flash-latest" in n
                else 3 if "1.5-flash" in n
                else 4
            )
            latest_bonus = 0 if n.endswith("-latest") else 1
            return (gen, latest_bonus, len(n))

        flash.sort(key=rank)
        names = [m.name.split("/")[-1] for m in flash]
        if names:
            return names
    except Exception:
        pass

    # Hard-coded fallback if list_models is unavailable.
    return [
        "gemini-2.0-flash",
        "gemini-2.5-flash",
        "gemini-flash-latest",
        "gemini-1.5-flash-latest",
    ]
