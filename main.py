"""
Development script for testing date_shift library functionality.

This script provides a quick end-to-end test of the date_shift library
with hardcoded test inputs.
"""

from typing import Any, Dict

from nuh_helper import shift_excel_dates


def main():
    """
    Main entry point with hardcoded inputs for dev/testing.
    """
    # Toggle between sample files as needed
    input_file = "test_unknown.xlsx"
    output_file = "test_shifted.xlsx"

    # Configuration: sheet name -> {patient_id_col, date_columns}
    # Currently based on test.xlsx structure:
    # - 'patients' sheet: patient_id, gender, dob (date column)
    # - 'labs' sheet: patient_id, test_date (date column), result
    sheet_configs: Dict[str, Dict[str, Any]] = {
        "patients": {
            "patient_id_col": "patient_id",
            "date_columns": ["dob", "date_of_diagnosis"],
            "header_row": 1,  # column names on second row (zero-based index 1)
        },
        "labs": {
            "patient_id_col": "patient_id",
            "date_columns": ["test_date"],
            "header_row": 1,
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
        patient_header_row=1,
        patient_skip_rows=None,
        date_format="YYYY-MM-DD",  # Format dates as YYYY-MM-DD
    )

    print("\nDate shifting complete!")
    print(f"Output file: {output_file}")


if __name__ == "__main__":
    main()
