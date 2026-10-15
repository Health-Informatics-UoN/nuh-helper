"""
Date shifting for patient data in Excel spreadsheets.

Consistently shifts dates for patient IDs across multiple sheets and columns
in an Excel file, with support for reproducible shifts using a linking table.
"""

from nuh_helper.date_shift.mappings import (  # noqa: E402
    generate_shift_mappings,
)
from nuh_helper.date_shift.stream_shift import (
    shift_excel_dates_inplace,
)

# Re-export public API
__all__ = [
    generate_shift_mappings,
    shift_excel_dates_inplace,
]
