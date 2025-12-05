"""
Date shifting tool for patient data in Excel spreadsheets.

This module provides functionality to consistently shift dates for patient IDs
across multiple sheets and columns in an Excel file, with support for
reproducible shifts using a linking table.
"""

import random
from pathlib import Path
from typing import Any, Dict, List, Optional, cast

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


def apply_date_shifts(
    df: pd.DataFrame,
    patient_id_col: str,
    date_columns: List[str],
    shift_mappings: pd.DataFrame,
) -> pd.DataFrame:
    """
    Apply date shifts to specified columns in a DataFrame.

    Args:
        df: DataFrame containing patient data.
        patient_id_col: Name of the column containing patient IDs.
        date_columns: List of column names containing dates to shift.
        shift_mappings: DataFrame with 'patient_id' and 'shift_days' columns.

    Returns:
        DataFrame with shifted dates.
    """
    df = df.copy()
    shift_dict = dict(
        zip(shift_mappings["patient_id"], shift_mappings["shift_days"])
    )

    for date_col in date_columns:
        if date_col not in df.columns:
            continue

        # Convert to datetime if not already
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

        # Apply shifts
        df[date_col] = df.apply(
            lambda row: (
                row[date_col] + pd.Timedelta(days=shift_dict.get(row[patient_id_col], 0))
                if pd.notna(row[date_col]) and row[patient_id_col] in shift_dict
                else row[date_col]
            ),
            axis=1,
        )

        # Convert back to date-only format (removes time component)
        # Convert datetime to Python date objects (Excel will display as date-only)
        mask = df[date_col].notna()
        df.loc[mask, date_col] = df.loc[mask, date_col].apply(lambda x: x.date())
        df.loc[~mask, date_col] = None

    return df


def shift_excel_dates(
    input_file: str,
    output_file: str,
    patient_sheet: str,
    patient_id_col: str,
    sheet_configs: Dict[str, Dict[str, str]],
    min_shift_days: int = -15,
    max_shift_days: int = 15,
    linking_table_path: Optional[str] = None,
    linking_table_output: Optional[str] = None,
    seed: Optional[int] = None,
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
        min_shift_days: Minimum number of days to shift (default: -15).
        max_shift_days: Maximum number of days to shift (default: 15).
        linking_table_path: Optional path to existing linking table CSV for reproducibility.
        linking_table_output: Path to save the linking table CSV (default: 'shift_mappings.csv').
        seed: Optional random seed for generating shifts.
    """
    # Read patient IDs from the central patient sheet
    patient_df = pd.read_excel(input_file, sheet_name=patient_sheet)
    if patient_id_col not in patient_df.columns:
        raise ValueError(
            f"Patient ID column '{patient_id_col}' not found in sheet '{patient_sheet}'"
        )

    patient_ids = patient_df[patient_id_col].dropna().unique().tolist()

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

    # Process each sheet
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        excel_file = pd.ExcelFile(input_file, engine="openpyxl")

        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)

            # Check if this sheet needs date shifting
            if sheet_name in sheet_configs:
                config = sheet_configs[cast(str, sheet_name)]
                sheet_patient_id_col: str = cast(str, config["patient_id_col"])
                date_columns: List[str] = cast(List[str], config["date_columns"])

                if sheet_patient_id_col not in df.columns:
                    raise ValueError(
                        f"Patient ID column '{sheet_patient_id_col}' not found in sheet '{sheet_name}'"
                    )

                df = apply_date_shifts(
                    df, sheet_patient_id_col, date_columns, shift_mappings
                )

            # Write the sheet (shifted or unshifted)
            df.to_excel(writer, sheet_name=cast(str, sheet_name), index=False)

    # Save linking table
    if linking_table_output:
        shift_mappings.to_csv(linking_table_output, index=False)
    else:
        shift_mappings.to_csv("shift_mappings.csv", index=False)


def main():
    """
    Main entry point with hardcoded inputs for dev/testing.
    """
    input_file = "test.xlsx"
    output_file = "test_shifted.xlsx"

    # Configuration: sheet name -> {patient_id_col, date_columns}
    # Currently based on test.xlsx structure:
    # - 'patients' sheet: patient_id, gender, dob (date column)
    # - 'labs' sheet: patient_id, test_date (date column), result
    sheet_configs: Dict[str, Dict[str, Any]] = {
        "patients": {
            "patient_id_col": "patient_id",
            "date_columns": ["dob"],
        },
        "labs": {
            "patient_id_col": "patient_id",
            "date_columns": ["test_date"],
        },
    }

    # Run the date shifting
    shift_excel_dates(
        input_file=input_file,
        output_file=output_file,
        patient_sheet="patients",
        patient_id_col="patient_id",
        sheet_configs=sheet_configs,
        min_shift_days=-15,
        max_shift_days=15,
        linking_table_path=None,  # Optional: path to existing linking table
        linking_table_output="shift_mappings.csv",
        seed=42,  # Optional: for reproducibility
    )

    print("\nDate shifting complete!")
    print(f"Output file: {output_file}")


if __name__ == "__main__":
    main()
