"""
Helper functions for enabling studies.

This library provides utilities for study enablement, including:
- **Date shifting**: consistently shift dates for patient IDs in Excel/DataFrames
- **Dataset profiling**: profile a dataset into a Scan Report
"""

import logging

logging.getLogger(__name__).addHandler(logging.NullHandler())

from nuh_helper.date_shift import (  # noqa: E402
    apply_date_shifts,
    generate_shift_mappings,
    load_shift_mappings,
    shift_excel_dates,
)
from nuh_helper.profile import generate_scan_report  # noqa: E402

__all__ = [
    "shift_excel_dates",
    "apply_date_shifts",
    "generate_shift_mappings",
    "load_shift_mappings",
    "generate_scan_report",
]
