"""Shift mapping generation and loading for date-shift."""

import logging
import random

import pandas as pd

logger = logging.getLogger(__name__)


def generate_shift_mappings(
    patient_ids: list[str],
    min_shift_days: int = -15,
    max_shift_days: int = 15,
    seed: int | None = None,
) -> pd.DataFrame:
    """
    Generate random shift mappings for patient IDs.

    Args:
        patient_ids: List of patient IDs to generate shifts for.
        min_shift_days: Minimum number of days to shift (default: -15).
        max_shift_days: Maximum number of days to shift (default: 15).
        seed: Optional random seed for reproducibility.

    Returns:
        DataFrame with columns 'patient_id' and 'shift_days'.
    """
    if seed is not None:
        random.seed(seed)

    shifts = [random.randint(min_shift_days, max_shift_days) for _ in patient_ids]
    return pd.DataFrame({"patient_id": patient_ids, "shift_days": shifts})


def load_shift_mappings(csv_path: str) -> pd.DataFrame:
    """
    Load shift mappings from a CSV file.

    Args:
        csv_path: Path to the CSV file containing shift mappings.
                  Expected columns: 'patient_id' and 'shift_days'.

    Returns:
        DataFrame with shift mappings.
    """
    df = pd.read_csv(csv_path)
    if "patient_id" not in df.columns or "shift_days" not in df.columns:
        raise ValueError("CSV must contain 'patient_id' and 'shift_days' columns")
    logger.info("Loaded %d shift mapping(s) from '%s'", len(df), csv_path)
    return df
