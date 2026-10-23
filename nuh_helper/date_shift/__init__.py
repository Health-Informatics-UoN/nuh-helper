"""
Date shifting for patient data in Excel spreadsheets.

Consistently shifts dates for patient IDs across multiple sheets and columns
in an Excel file, with support for reproducible shifts using a linking table.
"""

from nuh_helper.date_shift.stream_shift import (
    shift_excel_dates_inplace,
)

# Re-export public API
__all__ = [
    shift_excel_dates_inplace,
]
