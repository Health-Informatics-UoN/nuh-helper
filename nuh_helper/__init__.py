"""
Helper functions for enabling studies.

This library provides utilities for study enablement, including:
- **Date shifting**: consistently shift dates for patient IDs in Excel/DataFrames
- **Dataset profiling**: profile a dataset into a Scan Report
"""

import logging

logging.getLogger(__name__).addHandler(logging.NullHandler())

from nuh_helper.date_shift import (  # noqa: E402
    shift_excel_dates_inplace,
)
from nuh_helper.profile import generate_scan_report  # noqa: E402

__all__ = [
    generate_scan_report,
    shift_excel_dates_inplace,
]
