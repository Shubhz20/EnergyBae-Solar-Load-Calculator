"""
openpyxl-based writer for the Energybae Solar Load Excel Template.

Design rules:
- We NEVER touch cells that contain formulas.
- We only write to a configurable map of named input cells.
- The workbook is loaded with keep_vba=False and data_only=False so all
  existing formulas, styles, and named ranges remain intact.
"""

import io
from typing import Any, Dict, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string


# ---------------------------------------------------------------------------
# Default cell map for the Solar Load Template.
#
# These coordinates target the typical "Customer Details / Load Inputs" panel
# at the top of the Energybae Solar Load template. Adjust if the template
# layout changes - this is the ONLY place that needs editing.
# ---------------------------------------------------------------------------
DEFAULT_CELL_MAP: Dict[str, str] = {
    "consumer_name":           "C5",
    "consumer_number":         "C6",
    "billing_unit":            "C7",
    "tariff_category":         "C8",
    "connected_load_kw":       "C9",
    "avg_monthly_consumption": "C10",
}

# Optional: restrict writes to a specific sheet. None -> active sheet.
DEFAULT_SHEET_NAME: Optional[str] = None


def _to_number(value: Any) -> Any:
    """Convert numeric-looking strings to int/float so Excel treats them as numbers."""
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return value
    s = str(value).strip().replace(",", "")
    try:
        if "." in s:
            return float(s)
        return int(s)
    except ValueError:
        return value  # leave as text


def _is_formula_cell(cell) -> bool:
    v = cell.value
    return isinstance(v, str) and v.startswith("=")


def _split_coord(coord: str) -> Tuple[str, int]:
    col, row = coordinate_from_string(coord)
    return col, row


def fill_solar_load_template(
    template_bytes: bytes,
    data: Dict[str, Any],
    cell_map: Optional[Dict[str, str]] = None,
    sheet_name: Optional[str] = DEFAULT_SHEET_NAME,
) -> bytes:
    """
    Write `data` into the input cells of the Solar Load template.

    Parameters
    ----------
    template_bytes : bytes
        Raw bytes of the .xlsx template.
    data : dict
        Verified field values from the UI.
    cell_map : dict, optional
        Override the default field -> cell mapping.
    sheet_name : str, optional
        Worksheet to write into. Defaults to the active sheet.

    Returns
    -------
    bytes : the filled .xlsx as a byte string ready for download.
    """
    cell_map = cell_map or DEFAULT_CELL_MAP

    wb = load_workbook(filename=io.BytesIO(template_bytes), keep_vba=False)

    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.active

    numeric_fields = {"connected_load_kw", "avg_monthly_consumption"}

    for field, coord in cell_map.items():
        if field not in data:
            continue
        value = data.get(field)
        if value is None or value == "":
            continue

        cell = ws[coord]

        # Hard rule: never overwrite a formula cell.
        if _is_formula_cell(cell):
            continue

        if field in numeric_fields:
            value = _to_number(value)

        cell.value = value

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
