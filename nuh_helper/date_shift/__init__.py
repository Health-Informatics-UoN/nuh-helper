"""
Date shifting for patient data in Excel spreadsheets.

Consistently shifts dates for patient IDs across multiple sheets and columns
in an Excel file, with support for reproducible shifts using a linking table.
"""

import random
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, cast

import pandas as pd


def generate_shift_mappings(
    patient_ids: List[str],
    min_shift_days: int = -15,
    max_shift_days: int = 15,
    seed: Optional[int] = None,
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

    shifts = [
        random.randint(min_shift_days, max_shift_days) for _ in patient_ids
    ]
    return pd.DataFrame(
        {"patient_id": patient_ids, "shift_days": shifts}
    )


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
        raise ValueError(
            "CSV must contain 'patient_id' and 'shift_days' columns"
        )
    return df


def _parse_date_value(value: Any) -> Optional[pd.Timestamp]:
    """Parse a value into a pandas Timestamp if possible."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None

    # Already datetime-like
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return pd.to_datetime(value, errors="coerce")

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
            except Exception:
                pass

        # Fallback: let pandas try with dayfirst to handle ambiguous strings
        parsed = pd.to_datetime(v, errors="coerce", dayfirst=True, infer_datetime_format=True)
        return parsed if pd.notna(parsed) else None

    # Anything else: no parse
    return None


def _normalize_patient_id(value: Any) -> Optional[str]:
    """Normalize patient IDs by stripping whitespace and converting to string."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None

    if isinstance(value, str):
        v = value.strip()
        return v if v else None

    v = str(value).strip()
    return v if v else None


def apply_date_shifts(
    df: pd.DataFrame,
    patient_id_col: str,
    date_columns: List[str],
    shift_mappings: pd.DataFrame,
    date_format: Optional[str] = None,
) -> pd.DataFrame:
    """
    Apply date shifts to specified columns in a DataFrame.

    Args:
        df: DataFrame containing patient data.
        patient_id_col: Name of the column containing patient IDs.
        date_columns: List of column names containing dates to shift.
        shift_mappings: DataFrame with 'patient_id' and 'shift_days' columns.
        date_format: Optional date format string (e.g., 'YYYY-MM-DD').
                     Note: This parameter is kept for API compatibility but formatting
                     is applied at the Excel cell level, not in the DataFrame.

    Returns:
        DataFrame with shifted dates.
    """
    df = df.copy()

    # Normalize patient IDs in the working DataFrame to align with mapping keys
    df[patient_id_col] = df[patient_id_col].apply(_normalize_patient_id)

    shift_dict = dict(
        zip(shift_mappings["patient_id"], shift_mappings["shift_days"])
    )

    for date_col in date_columns:
        if date_col not in df.columns:
            continue

        # Parse flexible date strings (handles YYYY-DD-MM and placeholders like "Unknown")
        df[date_col] = df[date_col].apply(_parse_date_value)

        # Apply shifts
        df[date_col] = df.apply(
            lambda row: (
                row[date_col] + pd.Timedelta(days=shift_dict.get(row[patient_id_col], 0))
                if row[date_col] is not None and row[patient_id_col] in shift_dict
                else row[date_col]
            ),
            axis=1,
        )

        # Convert back to date-only format (removes time component)
        df[date_col] = df[date_col].apply(
            lambda x: x.date() if isinstance(x, (pd.Timestamp, datetime, date)) else None
        )

    return df


def shift_excel_dates(
    input_file: str,
    output_file: str,
    patient_sheet: str,
    patient_id_col: str,
    sheet_configs: Dict[str, Dict[str, Any]],
    min_shift_days: int = -15,
    max_shift_days: int = 15,
    linking_table_path: Optional[str] = None,
    linking_table_output: Optional[str] = None,
    seed: Optional[int] = None,
    patient_header_row: int = 0,
    patient_skip_rows: Optional[List[int]] = None,
    date_format: Optional[str] = None,
) -> None:
    """
    Shift dates in an Excel file for patient IDs consistently across sheets.

    Args:
        input_file: Path to input Excel file.
        output_file: Path to output Excel file with shifted dates.
        patient_sheet: Name of the sheet containing patient IDs.
        patient_id_col: Name of the column containing patient IDs in the patient sheet.
        sheet_configs: Dictionary mapping sheet names to configuration dicts.
                      Each config dict should have:
                      - 'patient_id_col': Name of patient ID column in that sheet
                      - 'date_columns': List of date column names to shift
                      Optional per-sheet header handling:
                      - 'header_row': zero-based row index of the column names (default 0)
                      - 'skip_rows': list of zero-based row indices to skip (e.g. description rows)
        min_shift_days: Minimum number of days to shift (default: -15).
        max_shift_days: Maximum number of days to shift (default: 15).
        linking_table_path: Optional path to existing linking table CSV for reproducibility.
        linking_table_output: Path to save the linking table CSV (default: 'shift_mappings.csv').
        seed: Optional random seed for generating shifts.
        patient_header_row: Zero-based header row index for the patient sheet (default: 0).
        patient_skip_rows: Optional rows to skip when reading the patient sheet (e.g. description rows).
        date_format: Optional Excel date format string (e.g., 'YYYY-MM-DD', 'yyyy-mm-dd').
                     If None, Excel's default date format is used.
                     Common formats: 'YYYY-MM-DD', 'MM/DD/YYYY', 'DD-MM-YYYY', etc.
    """
    def _read_sheet_with_structure(
        excel_file: pd.ExcelFile,
        sheet_name: str,
        header_row: int = 0,
    ) -> Tuple[pd.DataFrame, pd.DataFrame, List[List[Any]]]:
        """
        Read a sheet preserving description rows and structure.

        Returns:
            Tuple of (data_df, description_df, description_rows)
            - data_df: DataFrame with header row as column names and data rows
            - description_df: DataFrame with description rows (if any)
            - description_rows: List of description row data (for writing back)
        """
        # Read entire sheet without header to preserve all rows
        full_df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

        if header_row == 0:
            # No description rows, header is first row
            description_rows: List[List[Any]] = []
            description_df = pd.DataFrame()
            # Use first row as header
            data_df = pd.read_excel(
                excel_file, sheet_name=sheet_name, header=0
            )
        else:
            # Extract description rows (rows before header_row)
            description_rows = full_df.iloc[:header_row].values.tolist()
            description_df = full_df.iloc[:header_row].copy()

            # Read data with header_row as column names
            data_df = pd.read_excel(
                excel_file, sheet_name=sheet_name, header=header_row
            )

        return data_df, description_df, description_rows

    def _write_sheet_with_structure(
        writer: pd.ExcelWriter,
        sheet_name: str,
        data_df: pd.DataFrame,
        description_rows: List[List[Any]],
        header_row: int,
        date_columns: Optional[List[str]] = None,
        date_format: Optional[str] = None,
    ) -> None:
        """
        Write a sheet preserving description rows and structure.

        Args:
            writer: ExcelWriter instance.
            sheet_name: Name of the sheet to write.
            data_df: DataFrame with data to write.
            description_rows: List of description rows to write at the top.
            header_row: Row index where header should be written.
            date_columns: Optional list of date column names to format.
            date_format: Optional Excel date format string (e.g., 'yyyy-mm-dd').
        """
        # Convert Python date format to Excel format
        excel_date_format = None
        if date_format:
            # Convert common Python format strings to Excel format
            # Excel uses lowercase: yyyy (year), mm (month), dd (day)
            excel_date_format = (
                date_format.replace('YYYY', 'yyyy')
                          .replace('YY', 'yy')
                          .replace('DD', 'dd')
                          .replace('MM', 'mm')
            )

        # Calculate where to start writing data (after description rows + header row)
        data_start_row = len(description_rows) + 1 if description_rows else 1

        # Write data without header first (header=False), then we'll add header manually
        data_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            header=False,
            startrow=data_start_row
        )

        # Get the workbook and worksheet
        workbook = writer.book
        worksheet = workbook[sheet_name]

        # Write description rows at the top
        if description_rows:
            for row_idx, desc_row in enumerate(description_rows, start=1):
                for col_idx, value in enumerate(desc_row, start=1):
                    cell_value = value
                    # Handle NaN values
                    if pd.isna(cell_value):
                        cell_value = None
                    worksheet.cell(row=row_idx, column=col_idx, value=cell_value)

        # Write header row (after description rows)
        header_row_idx = len(description_rows) + 1 if description_rows else 1
        for col_idx, col_name in enumerate(data_df.columns, start=1):
            worksheet.cell(row=header_row_idx, column=col_idx, value=col_name)

        # Apply date formatting to date columns if specified
        if excel_date_format and date_columns:
            # Find column indices for date columns
            col_map = {col: idx + 1 for idx, col in enumerate(data_df.columns)}
            for date_col in date_columns:
                if date_col in col_map:
                    col_idx = col_map[date_col]
                    # Apply format to all data rows (skip description and header rows)
                    for row_idx in range(data_start_row + 1, data_start_row + len(data_df) + 1):
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        # Only format if cell has a date value
                        if cell.value is not None:
                            cell.number_format = excel_date_format

    # Read patient IDs from the central patient sheet
    patient_excel = pd.ExcelFile(input_file, engine="openpyxl")
    patient_df, _, _ = _read_sheet_with_structure(
        patient_excel,
        sheet_name=patient_sheet,
        header_row=patient_header_row,
    )
    if patient_id_col not in patient_df.columns:
        raise ValueError(
            f"Patient ID column '{patient_id_col}' not found in sheet '{patient_sheet}'"
        )

    patient_ids = (
        patient_df[patient_id_col]
        .apply(_normalize_patient_id)
        .dropna()
        .unique()
        .tolist()
    )

    # Generate or load shift mappings
    if linking_table_path and Path(linking_table_path).exists():
        shift_mappings = load_shift_mappings(linking_table_path)
        # Filter to only include patient IDs that exist in the data
        shift_mappings = shift_mappings[
            shift_mappings["patient_id"].isin(patient_ids)
        ]
        # Add any missing patient IDs with random shifts
        existing_ids = set(shift_mappings["patient_id"])
        missing_ids = [pid for pid in patient_ids if pid not in existing_ids]
        if missing_ids:
            new_shifts = generate_shift_mappings(
                missing_ids, min_shift_days, max_shift_days, seed
            )
            shift_mappings = pd.concat([shift_mappings, new_shifts], ignore_index=True)
    else:
        shift_mappings = generate_shift_mappings(
            patient_ids, min_shift_days, max_shift_days, seed
        )

    # Normalize and deduplicate patient IDs in the mapping
    shift_mappings["patient_id"] = shift_mappings["patient_id"].apply(_normalize_patient_id)
    shift_mappings = shift_mappings.dropna(subset=["patient_id"]).drop_duplicates(subset=["patient_id"], keep="first")

    # Process each sheet
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        excel_file = pd.ExcelFile(input_file, engine="openpyxl")

        for sheet_name in excel_file.sheet_names:
            # Default header handling unless overridden per sheet
            default_header_row = 0

            header_row = default_header_row

            # Track date columns for formatting
            sheet_date_columns: Optional[List[str]] = None

            # Check if this sheet needs date shifting
            if sheet_name in sheet_configs:
                config = sheet_configs[cast(str, sheet_name)]
                sheet_patient_id_col: str = cast(str, config["patient_id_col"])
                date_columns: List[str] = cast(List[str], config["date_columns"])
                header_row = cast(int, config.get("header_row", header_row))
                sheet_date_columns = date_columns

            # Read sheet preserving structure
            df, description_df, description_rows = _read_sheet_with_structure(
                excel_file,
                sheet_name=cast(str, sheet_name),
                header_row=header_row,
            )

            if sheet_name in sheet_configs:
                if sheet_patient_id_col not in df.columns:
                    raise ValueError(
                        f"Patient ID column '{sheet_patient_id_col}' not found in sheet '{sheet_name}'"
                    )

                df = apply_date_shifts(
                    df, sheet_patient_id_col, date_columns, shift_mappings, date_format=None
                )

            # Write the sheet preserving description rows (shifted or unshifted)
            _write_sheet_with_structure(
                writer,
                sheet_name=cast(str, sheet_name),
                data_df=df,
                description_rows=description_rows,
                header_row=header_row,
                date_columns=sheet_date_columns,
                date_format=date_format,
            )

    # Save linking table
    if linking_table_output:
        shift_mappings.to_csv(linking_table_output, index=False)
    else:
        shift_mappings.to_csv("shift_mappings.csv", index=False)


__all__ = [
    "shift_excel_dates",
    "apply_date_shifts",
    "generate_shift_mappings",
    "load_shift_mappings",
]
