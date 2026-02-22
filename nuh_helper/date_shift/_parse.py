"""Parsing and normalization helpers for date-shift."""

from datetime import date, datetime

import pandas as pd


def _parse_date_value(
    value: object,
) -> pd.Timestamp | None:
    """Parse a value into a pandas Timestamp if possible."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None

    # Already datetime-like
    if isinstance(value, (pd.Timestamp, datetime, date)):
        result = pd.to_datetime(value, errors="coerce")
        return result if pd.notna(result) else None

    if isinstance(value, str):
        v = value.strip()
        if not v or v.lower() in {"unknown", "unk", "unkown", "n/a", "none", "null"}:
            return None

        # Try a handful of common formats, including YYYY-DD-MM found in some feeds
        for fmt in ("%Y-%m-%d", "%Y-%d-%m", "%d-%m-%Y", "%m-%d-%Y"):
            try:
                parsed = pd.to_datetime(v, format=fmt, errors="coerce")
                if pd.notna(parsed):
                    return parsed
            except (TypeError, ValueError):
                # If parsing with this specific format fails, silently try the next format.
                pass

        # Fallback: let pandas try with dayfirst to handle ambiguous strings
        parsed = pd.to_datetime(v, errors="coerce", dayfirst=True)
        return parsed if pd.notna(parsed) else None

    # Anything else: no parse
    return None


def _normalize_patient_id(value: object) -> str | None:
    """Normalize patient IDs by stripping whitespace and converting to string."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None

    if isinstance(value, str):
        v = value.strip()
        return v if v else None

    v = str(value).strip()
    return v if v else None
