"""
Gemini 1.5 Flash extractor for MSEDCL electricity bills.

Returns a JSON object with:
    consumer_name, consumer_number, billing_unit,
    connected_load_kw, tariff_category, avg_monthly_consumption
"""

import json
import mimetypes
import re
from pathlib import Path
from typing import Any, Dict

import google.generativeai as genai


MODEL_NAME = "gemini-1.5-flash"

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
    Send the bill (PDF or image) to Gemini 1.5 Flash and return the parsed
    field dictionary.
    """
    if not api_key:
        raise ValueError("Gemini API key is required.")

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(MODEL_NAME)

    file_part = _load_file_part(file_path)

    response = model.generate_content(
        [PROMPT, file_part],
        generation_config={
            "temperature": 0.0,
            "response_mime_type": "application/json",
        },
    )

    raw = response.text or ""
    data = _coerce_json(raw)
    return _normalise(data)
