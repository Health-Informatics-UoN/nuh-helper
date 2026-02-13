"""
Helper functions for enabling studies.

This library provides utilities for study enablement, including:
- **Date shifting**: consistently shift dates for patient IDs in Excel/DataFrames
- **Dataset profiling**: profile a dataset into a Scan Report
"""

from nuh_helper.date_shift import (
    apply_date_shifts,
    generate_shift_mappings,
    load_shift_mappings,
    shift_excel_dates,
)
from nuh_helper.profile import generate_scan_report

__all__ = [
    "shift_excel_dates",
    "apply_date_shifts",
    "generate_shift_mappings",
    "load_shift_mappings",
    "generate_scan_report",
]
