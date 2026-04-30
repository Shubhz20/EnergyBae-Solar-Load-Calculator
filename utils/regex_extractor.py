"""
Offline MSEDCL bill parser - zero API calls, zero quota.

Used as a fallback when Gemini quota is exhausted. Works on PDF bills
because they have selectable text. Image bills (scanned/photo) need OCR
(out of scope for the offline path - Gemini handles those).

Field heuristics tuned to common MSEDCL invoice layouts.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    import fitz  # PyMuPDF
except ImportError:  # pragma: no cover
    fitz = None  # type: ignore


HP_TO_KW = 0.7457


def extract_text_from_pdf(file_path: str) -> str:
    """Pull selectable text out of a PDF using PyMuPDF."""
    if fitz is None:
        raise RuntimeError(
            "PyMuPDF is not installed. Run `pip install pymupdf` to enable "
            "the offline MSEDCL extractor."
        )
    doc = fitz.open(file_path)
    try:
        return "\n".join(page.get_text() for page in doc)
    finally:
        doc.close()


def _first_match(text: str, patterns: List[str], flags: int = re.IGNORECASE) -> Optional[str]:
    for p in patterns:
        m = re.search(p, text, flags)
        if m:
            return m.group(1).strip()
    return None


def _to_float(s: Optional[str]) -> Optional[float]:
    if not s:
        return None
    cleaned = s.replace(",", "").strip()
    m = re.search(r"-?\d+(?:\.\d+)?", cleaned)
    return float(m.group(0)) if m else None


def _consumer_number(text: str) -> Optional[str]:
    # MSEDCL consumer numbers are 12 digits.
    patterns = [
        r"Consumer\s*(?:No\.?|Number)[^\d]{0,15}(\d{9,12})",
        r"Cons(?:umer)?\s*No\.?\s*[:\-]?\s*(\d{9,12})",
    ]
    n = _first_match(text, patterns)
    if n:
        return n
    # Last resort: any standalone 12-digit run.
    m = re.search(r"\b(\d{12})\b", text)
    return m.group(1) if m else None


def _consumer_name(text: str) -> Optional[str]:
    patterns = [
        r"Consumer\s*Name\s*[:\-]?\s*([A-Z][A-Z .,&'\-]{2,80})",
        r"Name\s*&?\s*Address\s*[:\-]?\s*([A-Z][A-Z .,&'\-]{2,80})",
        r"Bill\s*To\s*[:\-]?\s*([A-Z][A-Z .,&'\-]{2,80})",
    ]
    n = _first_match(text, patterns)
    if n:
        # Trim trailing newline-bound noise like "Address" or single chars.
        n = re.split(r"\s{2,}|\n", n)[0]
        return n.strip(" .,-")
    return None


def _billing_unit(text: str) -> Optional[str]:
    patterns = [
        r"B\.?U\.?\s*(?:Name|No\.?|Code)?\s*[:\-]?\s*([A-Z0-9\- ]{2,40})",
        r"Billing\s*Unit\s*[:\-]?\s*([A-Z0-9\- ]{2,40})",
    ]
    n = _first_match(text, patterns)
    if n:
        n = re.split(r"\s{2,}|\n", n)[0]
        return n.strip()
    return None


def _tariff_category(text: str) -> Optional[str]:
    patterns = [
        r"Tariff\s*(?:Category|Code)?\s*[:\-]?\s*([A-Z]{2}-[A-Z0-9]+(?:\s+[A-Za-z]+)?)",
        r"\b(LT-?[IVX0-9]+(?:\s+[A-Za-z]+)?)\b",
        r"\b(HT-?[IVX0-9]+(?:\s+[A-Za-z]+)?)\b",
    ]
    return _first_match(text, patterns)


def _connected_load_kw(text: str) -> Optional[float]:
    # Try kW first, then HP -> kW conversion, then kVA (treat ~ kW).
    m = re.search(
        r"Connected\s*Load\s*[:\-]?\s*([\d.,]+)\s*(KW|KVA|HP)",
        text,
        flags=re.IGNORECASE,
    )
    if m:
        value = _to_float(m.group(1))
        unit = m.group(2).upper()
        if value is None:
            return None
        if unit == "HP":
            return round(value * HP_TO_KW, 3)
        return value

    m = re.search(r"Sanctioned\s*Load\s*[:\-]?\s*([\d.,]+)\s*(KW|KVA|HP)", text, flags=re.IGNORECASE)
    if m:
        value = _to_float(m.group(1))
        unit = m.group(2).upper()
        if value is None:
            return None
        if unit == "HP":
            return round(value * HP_TO_KW, 3)
        return value

    return None


def _avg_monthly_consumption(text: str) -> Optional[float]:
    """
    Strategy:
      1. If a 6-month consumption history is present (numbers near
         "Consumption History" or "Last 6 Months"), average the kWh values.
      2. Otherwise pull the current month's "Units Consumed".
    """
    # 6-month history table
    block = re.search(
        r"(?:Consumption\s*History|Last\s*6\s*Months|Previous\s*Consumption)"
        r"(.{0,800})",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if block:
        nums = re.findall(r"\b(\d{2,5})\b", block.group(1))
        floats = [
            float(n)
            for n in nums
            # exclude year-like values (1900-2100) and very small noise.
            if 10 <= int(n) <= 50000 and not (1900 <= int(n) <= 2100)
        ]
        if len(floats) >= 3:
            tail = floats[:6]
            return round(sum(tail) / len(tail), 2)

    # Current month units consumed
    m = re.search(
        r"(?:Units\s*Consumed|Total\s*Units|Consumption)\s*[:\-]?\s*([\d,]+)",
        text,
        flags=re.IGNORECASE,
    )
    if m:
        return _to_float(m.group(1))

    return None


def extract_from_pdf(file_path: str) -> Dict[str, Any]:
    """
    Run the full offline MSEDCL extraction pipeline.

    Returns the same field shape as the Gemini extractor, with None for
    anything not confidently parsed - the UI's verification step lets the
    user fill in the gaps.
    """
    suffix = Path(file_path).suffix.lower()
    if suffix != ".pdf":
        raise ValueError(
            "Offline extractor only supports PDF bills. For images/scans, "
            "the system needs Gemini (vision)."
        )

    text = extract_text_from_pdf(file_path)

    return {
        "consumer_name":           _consumer_name(text),
        "consumer_number":         _consumer_number(text),
        "billing_unit":            _billing_unit(text),
        "tariff_category":         _tariff_category(text),
        "connected_load_kw":       _connected_load_kw(text),
        "avg_monthly_consumption": _avg_monthly_consumption(text),
    }
