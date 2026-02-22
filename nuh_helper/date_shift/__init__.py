"""
Date shifting for patient data in Excel spreadsheets.

Consistently shifts dates for patient IDs across multiple sheets and columns
in an Excel file, with support for reproducible shifts using a linking table.
"""

import contextlib
import logging
import random
import shutil
from datetime import date, datetime
from pathlib import Path
from typing import Any, cast

import pandas as pd
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.worksheet import Worksheet

logger = logging.getLogger(__name__)


def _get_row_values_resolving_merged(
    sheet: Worksheet, row_1based: int, max_col: int
) -> list[Any]:
    """
    Return values for one row, resolving merged cells to the top-left cell value.

    openpyxl stores the value only in the first cell of a merged range;
    other cells are MergedCell and have no value. This walks the row and
    fills in the value from the merge range's top-left for each column.
    """
    result: list[Any] = []
    for col in range(1, max_col + 1):
        cell = sheet.cell(row=row_1based, column=col)
        if isinstance(cell, MergedCell):
            for merged_range in sheet.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    top_left = sheet.cell(
                        row=merged_range.min_row, column=merged_range.min_col
                    )
                    result.append(top_left.value)
                    break
            else:
                result.append(None)
        else:
            result.append(cell.value)
    return result


def _description_merged_ranges(
    sheet: Worksheet, num_description_rows: int
) -> list[str]:
    """Merged range refs (e.g. 'A1:D1') in the first num_description_rows."""
    result: list[str] = []
    for merged_range in sheet.merged_cells.ranges:
        if merged_range.max_row <= num_description_rows:
            result.append(str(merged_range))
    return result


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
            except Exception:
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


def apply_date_shifts(
    df: pd.DataFrame,
    patient_id_col: str,
    date_columns: list[str],
    shift_mappings: pd.DataFrame,
    date_format: str | None = None,
) -> pd.DataFrame:
    """
    Apply date shifts to specified columns in a DataFrame.

    Args:
        df: pd.DataFrame containing patient data.
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
        zip(
            shift_mappings["patient_id"],
            shift_mappings["shift_days"],
            strict=True,
        )
    )

    for date_col in date_columns:
        if date_col not in df.columns:
            logger.warning(
                "Date column '%s' not found in DataFrame, skipping", date_col
            )
            continue

        # Parse flexible date strings (handles YYYY-DD-MM and placeholders "Unknown")
        non_null_before = df[date_col].notna().sum()
        df[date_col] = df[date_col].apply(_parse_date_value)
        parse_failures = non_null_before - sum(x is not None for x in df[date_col])
        if parse_failures > 0:
            logger.debug(
                "Column '%s': %d value(s) could not be parsed as dates",
                date_col,
                parse_failures,
            )

        # Apply shifts
        df[date_col] = df.apply(
            lambda row: (
                row[date_col]  # noqa: B023
                + pd.Timedelta(days=shift_dict.get(row[patient_id_col], 0))
                if row[date_col] is not None and row[patient_id_col] in shift_dict  # noqa: B023
                else row[date_col]  # noqa: B023
            ),
            axis=1,
        )

        # Convert back to date-only format (removes time component)
        df[date_col] = df[date_col].apply(
            lambda x: (
                x.date() if isinstance(x, (pd.Timestamp, datetime, date)) else None
            )
        )

    return df


def _read_sheet_with_structure(
    excel_file: pd.ExcelFile,
    sheet_name: str,
    header_row: int = 0,
    input_file: str | None = None,
    skip_rows_after_header: list[int] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, list[list[Any]], list[str]]:
    """
    Read a sheet preserving description rows and structure.

    Uses openpyxl to resolve merged cells in the header row so column names
    are correct when the sheet has merged cells. Rows whose indices are in
    skip_rows_after_header (0-based) are excluded from the data (e.g. a
    data-type row immediately below the header).

    Returns:
        Tuple of (data_df, description_df, description_rows,
        description_merged_ranges)
        - data_df: DataFrame with header row as column names and data rows
        - description_df: DataFrame with description rows (if any)
        - description_rows: List of description row data (for writing back)
        - description_merged_ranges: Merged range refs to preserve when writing
    """
    # Read entire sheet without header to preserve all rows
    full_df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
    max_col = full_df.shape[1]
    description_merged_ranges: list[str] = []

    # Resolve header row (and optional description merged ranges) via openpyxl
    if input_file and max_col:
        wb = load_workbook(input_file, read_only=False, data_only=True)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            # Header row in 1-based openpyxl
            header_row_1based = header_row + 1
            header_values = _get_row_values_resolving_merged(
                ws, header_row_1based, max_col
            )
            # Pandas-friendly column names (no None)
            columns = []
            for i, v in enumerate(header_values):
                if v is None or (
                    isinstance(v, float) and pd.isna(cast("float", v))
                ):
                    columns.append(f"Unnamed: {i}")
                else:
                    columns.append(str(v).strip() or f"Unnamed: {i}")

            if header_row > 0:
                description_merged_ranges = _description_merged_ranges(
                    ws, header_row
                )

            # Description rows (rows before header)
            description_rows = (
                full_df.iloc[:header_row].values.tolist()
                if header_row > 0
                else []
            )
            # Data rows: everything after header, optionally excluding some rows
            data_block = full_df.iloc[header_row + 1 :]
            if skip_rows_after_header:
                # Drop by 0-based index (data_block index = original row)
                to_drop = [
                    i
                    for i in data_block.index
                    if i in skip_rows_after_header
                ]
                data_block = data_block.drop(index=to_drop, errors="ignore")

            data_df = pd.DataFrame(data_block.values, columns=columns)
            description_df = (
                full_df.iloc[:header_row].copy()
                if header_row > 0
                else pd.DataFrame()
            )
            wb.close()
            return (
                data_df,
                description_df,
                description_rows,
                description_merged_ranges,
            )

    # Fallback when no openpyxl path or empty sheet
    if header_row == 0:
        description_rows = []
        description_df = pd.DataFrame()
        data_df = pd.read_excel(
            excel_file, sheet_name=sheet_name, header=0
        )
        if skip_rows_after_header:
            data_df = data_df.drop(index=skip_rows_after_header, errors="ignore")
    else:
        description_rows = full_df.iloc[:header_row].values.tolist()
        description_df = full_df.iloc[:header_row].copy()
        data_df = pd.read_excel(
            excel_file, sheet_name=sheet_name, header=header_row
        )
        if skip_rows_after_header:
            data_df = data_df.drop(index=skip_rows_after_header, errors="ignore")

    return (
        data_df,
        description_df,
        description_rows,
        description_merged_ranges,
    )


def shift_excel_dates(
    input_file: str,
    output_file: str,
    patient_sheet: str,
    patient_id_col: str,
    sheet_configs: dict[str, dict[str, Any]],
    min_shift_days: int = -15,
    max_shift_days: int = 15,
    linking_table_path: str | None = None,
    linking_table_output: str | None = None,
    seed: int | None = None,
    patient_header_row: int = 0,
    patient_skip_rows: list[int] | None = None,
    date_format: str | None = None,
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
                      - 'skip_rows_after_header': list of zero-based row indices to exclude
                        from data (e.g. a data-type row immediately below the header)
        min_shift_days: Minimum number of days to shift (default: -15).
        max_shift_days: Maximum number of days to shift (default: 15).
        linking_table_path: Optional path to existing linking table CSV for reproducibility.
        linking_table_output: Path to save the linking table CSV (default: 'shift_mappings.csv').
        seed: Optional random seed for generating shifts.
        patient_header_row: Zero-based header row index for the patient sheet (default: 0).
        patient_skip_rows: Optional zero-based row indices to exclude from patient data
                     (e.g. a data-type row immediately below the header).
        date_format: Optional Excel date format string (e.g., 'YYYY-MM-DD', 'yyyy-mm-dd').
                     If None, Excel's default date format is used.
                     Common formats: 'YYYY-MM-DD', 'MM/DD/YYYY', 'DD-MM-YYYY', etc.
    """  # noqa: E501
    logger.info("Shifting dates: '%s' → '%s'", input_file, output_file)
    logger.debug(
        "Shift range: %d to %d days, seed=%s", min_shift_days, max_shift_days, seed
    )

    def _write_sheet_with_structure(
        writer: pd.ExcelWriter,
        sheet_name: str,
        data_df: pd.DataFrame,
        description_rows: list[list[Any]],
        header_row: int,
        date_columns: list[str] | None = None,
        date_format: str | None = None,
        description_merged_ranges: list[str] | None = None,
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
            description_merged_ranges: Optional merged cell refs to restore
                (e.g. 'A1:D1').
        """
        # Convert Python date format to Excel format
        excel_date_format = None
        if date_format:
            # Convert common Python format strings to Excel format
            # Excel uses lowercase: yyyy (year), mm (month), dd (day)
            excel_date_format = (
                date_format.replace("YYYY", "yyyy")
                .replace("YY", "yy")
                .replace("DD", "dd")
                .replace("MM", "mm")
            )

        # Calculate where to start writing data (after description rows + header row)
        data_start_row = len(description_rows) + 1 if description_rows else 1

        # Write data without header first (header=False), then we'll add header manually
        data_df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False,
            header=False,
            startrow=data_start_row,
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
            # Restore merged cells in the description area
            if description_merged_ranges:
                for range_ref in description_merged_ranges:
                    with contextlib.suppress(Exception):
                        worksheet.merge_cells(range_ref)

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
                    for row_idx in range(
                        data_start_row + 1, data_start_row + len(data_df) + 1
                    ):
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        # Only format if cell has a date value
                        if cell.value is not None:
                            cell.number_format = excel_date_format

    # Read patient IDs from the central patient sheet.
    # If the patient sheet is in sheet_configs, use its header_row and
    # skip_rows_after_header so layout is defined in one place.
    effective_patient_header_row = patient_header_row
    effective_patient_skip_rows = patient_skip_rows
    if patient_sheet in sheet_configs:
        cfg = sheet_configs[patient_sheet]
        effective_patient_header_row = cast(
            int, cfg.get("header_row", patient_header_row)
        )
        effective_patient_skip_rows = cfg.get(
            "skip_rows_after_header", patient_skip_rows
        )

    patient_excel = pd.ExcelFile(input_file, engine="openpyxl")
    patient_df, _, _, _ = _read_sheet_with_structure(
        patient_excel,
        sheet_name=patient_sheet,
        header_row=effective_patient_header_row,
        input_file=input_file,
        skip_rows_after_header=effective_patient_skip_rows,
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
    logger.info(
        "Found %d patient(s) in sheet '%s'", len(patient_ids), patient_sheet
    )

    # Generate or load shift mappings
    if linking_table_path and Path(linking_table_path).exists():
        logger.info("Loading shift mappings from '%s'", linking_table_path)
        shift_mappings = load_shift_mappings(linking_table_path)
        # Filter to only include patient IDs that exist in the data
        shift_mappings = shift_mappings[shift_mappings["patient_id"].isin(patient_ids)]
        # Add any missing patient IDs with random shifts
        existing_ids = set(shift_mappings["patient_id"])
        missing_ids = [pid for pid in patient_ids if pid not in existing_ids]
        if missing_ids:
            logger.warning(
                "%d patient(s) had no entry in the linking table; new shifts generated",
                len(missing_ids),
            )
            new_shifts = generate_shift_mappings(
                missing_ids, min_shift_days, max_shift_days, seed
            )
            shift_mappings = pd.concat([shift_mappings, new_shifts], ignore_index=True)
    else:
        logger.info("Generating shift mappings for %d patient(s)", len(patient_ids))
        shift_mappings = generate_shift_mappings(
            patient_ids, min_shift_days, max_shift_days, seed
        )

    # Normalize and deduplicate patient IDs in the mapping
    shift_mappings["patient_id"] = shift_mappings["patient_id"].apply(
        _normalize_patient_id
    )
    shift_mappings = shift_mappings.dropna(subset=["patient_id"]).drop_duplicates(
        subset=["patient_id"], keep="first"
    )

    # Process each sheet
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        excel_file = pd.ExcelFile(input_file, engine="openpyxl")

        for sheet_name in excel_file.sheet_names:
            # Default header handling unless overridden per sheet
            default_header_row = 0

            header_row = default_header_row

            # Track date columns for formatting
            sheet_date_columns: list[str] | None = None

            skip_rows_after_header: list[int] | None = None

            # Check if this sheet needs date shifting
            if sheet_name in sheet_configs:
                config = sheet_configs[cast(str, sheet_name)]
                sheet_patient_id_col: str = cast(str, config["patient_id_col"])
                date_columns: list[str] = cast(list[str], config["date_columns"])
                header_row = cast(int, config.get("header_row", header_row))
                skip_rows_after_header = config.get("skip_rows_after_header")
                sheet_date_columns = date_columns
                logger.info(
                    "Shifting %d date column(s) in sheet '%s'",
                    len(date_columns),
                    sheet_name,
                )

            # Read sheet preserving structure
            df, description_df, description_rows, description_merged_ranges = (
                _read_sheet_with_structure(
                    excel_file,
                    sheet_name=cast(str, sheet_name),
                    header_row=header_row,
                    input_file=input_file,
                    skip_rows_after_header=skip_rows_after_header,
                )
            )

            if sheet_name in sheet_configs:
                if sheet_patient_id_col not in df.columns:
                    raise ValueError(
                        f"Patient ID column '{sheet_patient_id_col}' not found in sheet '{sheet_name}'"  # noqa: E501
                    )

                df = apply_date_shifts(
                    df,
                    sheet_patient_id_col,
                    date_columns,
                    shift_mappings,
                    date_format=None,
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
                description_merged_ranges=description_merged_ranges,
            )

    logger.info("Output written to '%s'", output_file)

    # Save linking table
    linking_path = linking_table_output or "shift_mappings.csv"
    shift_mappings.to_csv(linking_path, index=False)
    logger.info("Linking table saved to '%s'", linking_path)


def shift_excel_dates_inplace(
    input_file: str,
    output_file: str,
    patient_sheet: str,
    patient_id_col: str,
    sheet_configs: dict[str, dict[str, Any]],
    min_shift_days: int = -15,
    max_shift_days: int = 15,
    linking_table_path: str | None = None,
    linking_table_output: str | None = None,
    seed: int | None = None,
    patient_header_row: int = 0,
    patient_skip_rows: list[int] | None = None,
) -> None:
    """
    Shift dates in an Excel file, preserving all cell formatting.

    Unlike shift_excel_dates(), this function copies the input file and then
    modifies date cells directly via openpyxl, so all formatting (cell styles,
    merged cells, column widths, conditional formatting, etc.) is preserved.

    Args:
        input_file: Path to input Excel file.
        output_file: Path for the output file (copy of input with shifted dates).
        patient_sheet: Name of the sheet containing patient IDs.
        patient_id_col: Name of the column containing patient IDs in the patient sheet.
        sheet_configs: Dictionary mapping sheet names to configuration dicts.
                      Each config dict should have:
                      - 'patient_id_col': Name of patient ID column in that sheet
                      - 'date_columns': List of date column names to shift
                      Optional per-sheet header handling:
                      - 'header_row': zero-based row index of the column names (default 0)
                      - 'skip_rows_after_header': list of zero-based row indices to exclude
                        from data (e.g. a data-type row immediately below the header)
        min_shift_days: Minimum number of days to shift (default: -15).
        max_shift_days: Maximum number of days to shift (default: 15).
        linking_table_path: Optional path to existing linking table CSV for reproducibility.
        linking_table_output: Path to save the linking table CSV (default: 'shift_mappings.csv').
        seed: Optional random seed for generating shifts.
        patient_header_row: Zero-based header row index for the patient sheet (default: 0).
        patient_skip_rows: Optional zero-based row indices to exclude from patient data.
    """  # noqa: E501
    logger.info("Shifting dates in-place: '%s' → '%s'", input_file, output_file)
    logger.debug(
        "Shift range: %d to %d days, seed=%s", min_shift_days, max_shift_days, seed
    )

    # Copy the input file first so all formatting is preserved
    shutil.copy2(input_file, output_file)

    # Determine patient sheet header settings (may be overridden by sheet_configs)
    effective_patient_header_row = patient_header_row
    effective_patient_skip_rows = patient_skip_rows
    if patient_sheet in sheet_configs:
        cfg = sheet_configs[patient_sheet]
        effective_patient_header_row = cast(
            int, cfg.get("header_row", patient_header_row)
        )
        effective_patient_skip_rows = cfg.get(
            "skip_rows_after_header", patient_skip_rows
        )

    # Read patient IDs via the module-level structure-aware helper
    patient_excel = pd.ExcelFile(input_file, engine="openpyxl")
    patient_df, _, _, _ = _read_sheet_with_structure(
        patient_excel,
        sheet_name=patient_sheet,
        header_row=effective_patient_header_row,
        input_file=input_file,
        skip_rows_after_header=effective_patient_skip_rows,
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
    logger.info(
        "Found %d patient(s) in sheet '%s'", len(patient_ids), patient_sheet
    )

    # Generate or load shift mappings (same logic as shift_excel_dates)
    if linking_table_path and Path(linking_table_path).exists():
        logger.info("Loading shift mappings from '%s'", linking_table_path)
        shift_mappings = load_shift_mappings(linking_table_path)
        shift_mappings = shift_mappings[shift_mappings["patient_id"].isin(patient_ids)]
        existing_ids = set(shift_mappings["patient_id"])
        missing_ids = [pid for pid in patient_ids if pid not in existing_ids]
        if missing_ids:
            logger.warning(
                "%d patient(s) had no entry in the linking table; new shifts generated",
                len(missing_ids),
            )
            new_shifts = generate_shift_mappings(
                missing_ids, min_shift_days, max_shift_days, seed
            )
            shift_mappings = pd.concat([shift_mappings, new_shifts], ignore_index=True)
    else:
        logger.info("Generating shift mappings for %d patient(s)", len(patient_ids))
        shift_mappings = generate_shift_mappings(
            patient_ids, min_shift_days, max_shift_days, seed
        )

    shift_mappings["patient_id"] = shift_mappings["patient_id"].apply(
        _normalize_patient_id
    )
    shift_mappings = shift_mappings.dropna(subset=["patient_id"]).drop_duplicates(
        subset=["patient_id"], keep="first"
    )

    shift_dict: dict[str, int] = dict(
        zip(
            shift_mappings["patient_id"],
            shift_mappings["shift_days"],
            strict=True,
        )
    )

    # Open the copied workbook and modify date cells directly.
    # keep_links=False drops external link definitions so Excel doesn't repair them.
    # Defined names are cleared for the same reason — openpyxl doesn't round-trip
    # all named range syntax faithfully, causing Excel to report repair errors.
    wb = load_workbook(output_file, keep_links=False)
    wb.defined_names.clear()

    for sheet_name, config in sheet_configs.items():
        if sheet_name not in wb.sheetnames:
            logger.warning("Sheet '%s' not found in workbook, skipping", sheet_name)
            continue

        ws = cast(Worksheet, wb[sheet_name])
        sheet_patient_id_col: str = cast(str, config["patient_id_col"])
        date_columns: list[str] = cast(list[str], config["date_columns"])
        header_row: int = cast(int, config.get("header_row", 0))
        skip_rows_after_header: list[int] | None = config.get("skip_rows_after_header")

        max_col = ws.max_column or 0
        if not max_col:
            continue

        # Resolve column names from the header row (handles merged cells)
        header_row_1based = header_row + 1
        header_values = _get_row_values_resolving_merged(ws, header_row_1based, max_col)
        col_index: dict[str, int] = {}
        for i, val in enumerate(header_values, start=1):
            if val is not None and str(val).strip():
                col_index[str(val).strip()] = i

        if sheet_patient_id_col not in col_index:
            raise ValueError(
                f"Patient ID column '{sheet_patient_id_col}' not found in sheet '{sheet_name}'"  # noqa: E501
            )

        pid_col_idx = col_index[sheet_patient_id_col]
        date_col_indices: dict[str, int] = {}
        for col in date_columns:
            if col in col_index:
                date_col_indices[col] = col_index[col]
            else:
                logger.warning(
                    "Date column '%s' not found in sheet '%s', skipping", col, sheet_name  # noqa: E501
                )

        if not date_col_indices:
            continue

        # Build the set of openpyxl row numbers (1-based) to skip
        # skip_rows_after_header uses 0-based pandas-style indices from sheet start
        skip_row_set: set[int] = (
            {idx + 1 for idx in skip_rows_after_header}
            if skip_rows_after_header
            else set()
        )

        logger.info(
            "Shifting %d date column(s) in sheet '%s'",
            len(date_col_indices),
            sheet_name,
        )

        data_start_1based = header_row + 2  # 1-based, one past the header row
        for row_idx in range(data_start_1based, (ws.max_row or 0) + 1):
            if row_idx in skip_row_set:
                continue

            pid_cell = ws.cell(row=row_idx, column=pid_col_idx)
            pid = _normalize_patient_id(pid_cell.value)
            shift_days = shift_dict.get(pid) if pid is not None else None

            for date_col_idx in date_col_indices.values():
                cell = ws.cell(row=row_idx, column=date_col_idx)
                original_value = cell.value
                parsed = _parse_date_value(original_value)

                if parsed is None:
                    # Clear unparseable strings (e.g. "unknown") to match existing behaviour  # noqa: E501
                    if original_value is not None and not isinstance(
                        original_value, (datetime, date)
                    ):
                        cell.value = None
                    continue

                if shift_days is None:
                    continue

                shifted = parsed + pd.Timedelta(days=shift_days)
                # Preserve date vs datetime type so Excel number format stays valid
                if isinstance(original_value, date) and not isinstance(
                    original_value, datetime
                ):
                    cell.value = cast(Any, shifted.to_pydatetime().date())
                else:
                    cell.value = cast(Any, shifted.to_pydatetime())

    wb.save(output_file)
    logger.info("Output written to '%s'", output_file)

    # Save linking table
    linking_path = linking_table_output or "shift_mappings.csv"
    shift_mappings.to_csv(linking_path, index=False)
    logger.info("Linking table saved to '%s'", linking_path)


__all__ = [
    "shift_excel_dates",
    "shift_excel_dates_inplace",
    "apply_date_shifts",
    "generate_shift_mappings",
    "load_shift_mappings",
]
