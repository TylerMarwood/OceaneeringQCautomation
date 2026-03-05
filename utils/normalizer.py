"""
Value normalisation for cross-sheet comparison.

All values are reduced to one of four canonical types before comparison:
  - None          → empty / null (NaN, None, "", "N/A", etc.)
  - pd.Timestamp  → any date-like value, time component stripped
  - int / float   → any numeric value
  - str           → everything else, whitespace-stripped

Two normalised values that are equal under Python's == operator are considered
a match.  Crucially this means:
  "01/01/2024" (string)  ==  Excel date serial  ==  datetime(2024,1,1)
  "1.0" (string)         ==  1.0 (float)         ==  1 (int)
  None                   ==  None  (both empty)
"""

from __future__ import annotations

import re
from datetime import date, datetime
from typing import Any

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Date format catalogue — tried in order for string → date parsing
# ---------------------------------------------------------------------------
_DATE_FORMATS: list[str] = [
    "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y",
    "%Y/%m/%d", "%d %b %Y", "%d %B %Y", "%b %d, %Y",
    "%d/%m/%y", "%m/%d/%y", "%d-%m-%y", "%Y%m%d",
]

# Pre-compiled regex: detects numeric strings like "1", "1.5", "1,234.56"
_NUMERIC_RE = re.compile(r"^-?[\d,]+(\.\d+)?$")


# ---------------------------------------------------------------------------
# Null detection
# ---------------------------------------------------------------------------

_NULL_STRINGS = {"", "nan", "none", "null", "n/a", "na", "#n/a", "-"}


def is_null(value: Any) -> bool:
    """Return True if *value* represents a missing / empty cell."""
    if value is None:
        return True
    if isinstance(value, float) and np.isnan(value):
        return True
    if isinstance(value, pd.Timestamp) and pd.isna(value):
        return True
    if isinstance(value, str) and value.strip().lower() in _NULL_STRINGS:
        return True
    return False


# ---------------------------------------------------------------------------
# Single-value normalisation
# ---------------------------------------------------------------------------

def _try_parse_date_string(s: str) -> pd.Timestamp | None:
    """Return a date-only pd.Timestamp if *s* matches a known date pattern."""
    s = s.strip()
    for fmt in _DATE_FORMATS:
        try:
            return pd.Timestamp(datetime.strptime(s, fmt)).normalize()
        except (ValueError, TypeError):
            pass
    # Last resort: let pandas try
    try:
        ts = pd.to_datetime(s, dayfirst=True)
        if not pd.isna(ts):
            return ts.normalize()
    except (ValueError, TypeError, OverflowError):
        pass
    return None


def normalize_value(value: Any) -> Any:
    """
    Reduce *value* to a canonical, comparable form.

    Returns None, a pd.Timestamp, an int/float, or a stripped string.
    """
    if is_null(value):
        return None

    # ----- datetime / date / Timestamp -----
    if isinstance(value, pd.Timestamp):
        return value.normalize()
    if isinstance(value, (datetime, date)):
        return pd.Timestamp(value).normalize()

    # ----- string -----
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped or stripped.lower() in _NULL_STRINGS:
            return None

        # Try date
        date_val = _try_parse_date_string(stripped)
        if date_val is not None:
            return date_val

        # Try numeric (strip thousand-separating commas)
        clean_num = stripped.replace(",", "")
        if _NUMERIC_RE.match(clean_num):
            try:
                f = float(clean_num)
                return int(f) if f == int(f) else f
            except (ValueError, OverflowError):
                pass

        return stripped

    # ----- int / numpy integer -----
    if isinstance(value, (int, np.integer)):
        return int(value)

    # ----- float / numpy float -----
    if isinstance(value, (float, np.floating)):
        if np.isnan(value):
            return None
        f = float(value)
        return int(f) if f == int(f) else round(f, 10)

    # Fallback — convert to string
    return str(value).strip()


# ---------------------------------------------------------------------------
# Series-level normalisation
# ---------------------------------------------------------------------------

def normalize_series(series: pd.Series) -> pd.Series:
    """Apply normalize_value element-wise to a pandas Series."""
    return series.apply(normalize_value)


# ---------------------------------------------------------------------------
# Match helper
# ---------------------------------------------------------------------------

def values_match(a: Any, b: Any) -> bool:
    """
    Return True if two normalised values are considered equal.

    Handles None/None (match), None/value (no match), Timestamp comparison,
    and standard equality for everything else.
    """
    a_null = is_null(a) or a is None
    b_null = is_null(b) or b is None

    if a_null and b_null:
        return True
    if a_null or b_null:
        return False

    # Both Timestamps: compare directly (already date-only)
    if isinstance(a, pd.Timestamp) and isinstance(b, pd.Timestamp):
        return a == b

    # Type mismatch after normalisation — convert both to string for safety
    if type(a) is not type(b):
        # Allow int vs float comparison
        if isinstance(a, (int, float)) and isinstance(b, (int, float)):
            return float(a) == float(b)
        return str(a) == str(b)

    return a == b
