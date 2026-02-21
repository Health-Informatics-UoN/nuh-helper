"""End-to-end tests for shift_excel_dates using tests/data/patients.xlsx."""

from pathlib import Path

import pandas as pd
import pytest

from nuh_helper import shift_excel_dates

DATA_DIR = Path(__file__).parent / "data"
INPUT_FILE = DATA_DIR / "patients.xlsx"

# Config matching the structure of patients.xlsx
SHEET_CONFIGS = {
    "patients": {
        "patient_id_col": "patient_id",
        "date_columns": ["dob", "last_alive"],
    },
    "results": {
        "patient_id_col": "patient_id",
        "date_columns": ["date_result"],
    },
}


def run_shift(tmp_path: Path, **kwargs) -> tuple[Path, Path]:
    """Call shift_excel_dates with defaults for e2e tests.

    Returns (output_xlsx_path, linking_table_csv_path).
    Always writes a linking table so tests don't pollute the working directory.
    """
    output = tmp_path / "output.xlsx"
    linking = tmp_path / "linking.csv"
    shift_excel_dates(
        str(INPUT_FILE),
        str(output),
        patient_sheet="patients",
        patient_id_col="patient_id",
        sheet_configs=SHEET_CONFIGS,
        linking_table_output=str(linking),
        **kwargs,
    )
    return output, linking


def read_sheet(path: Path, sheet: str) -> pd.DataFrame:
    return pd.read_excel(str(path), sheet_name=sheet)


class TestOutputStructure:
    def test_output_file_is_created(self, tmp_path):
        output, _ = run_shift(tmp_path, seed=42)
        assert output.exists()

    def test_linking_table_is_created_with_one_row_per_patient(self, tmp_path):
        _, linking = run_shift(tmp_path, seed=42)
        df = pd.read_csv(linking)
        assert set(df.columns) == {"patient_id", "shift_days"}
        assert len(df) == 5

    def test_output_sheets_match_input_sheets(self, tmp_path):
        output, _ = run_shift(tmp_path, seed=42)
        input_sheets = pd.ExcelFile(str(INPUT_FILE)).sheet_names
        output_sheets = pd.ExcelFile(str(output)).sheet_names
        assert output_sheets == input_sheets

    def test_non_date_columns_are_unchanged(self, tmp_path):
        output, _ = run_shift(tmp_path, seed=42)
        pd.testing.assert_series_equal(
            read_sheet(INPUT_FILE, "patients")["name"],
            read_sheet(output, "patients")["name"],
        )
        pd.testing.assert_series_equal(
            read_sheet(INPUT_FILE, "results")["measurement"],
            read_sheet(output, "results")["measurement"],
        )


class TestDateShifting:
    def test_dates_are_shifted_by_amounts_in_linking_table(self, tmp_path):
        output, linking = run_shift(tmp_path, seed=42)

        shift_dict = dict(zip(
            pd.read_csv(linking)["patient_id"],
            pd.read_csv(linking)["shift_days"],
        ))
        input_df = read_sheet(INPUT_FILE, "patients")
        output_df = read_sheet(output, "patients")

        for pid, expected_days in shift_dict.items():
            in_dob = pd.Timestamp(input_df.loc[input_df["patient_id"] == pid, "dob"].iloc[0])
            out_dob = pd.Timestamp(output_df.loc[output_df["patient_id"] == pid, "dob"].iloc[0])
            assert (out_dob - in_dob).days == expected_days, (
                f"Patient {pid}: expected shift of {expected_days} days, "
                f"got {(out_dob - in_dob).days}"
            )

    def test_shift_is_consistent_across_sheets(self, tmp_path):
        """Each patient's dates shift by the same number of days in every sheet."""
        output, linking = run_shift(tmp_path, seed=42)

        shift_dict = dict(zip(
            pd.read_csv(linking)["patient_id"],
            pd.read_csv(linking)["shift_days"],
        ))
        input_results = read_sheet(INPUT_FILE, "results")
        output_results = read_sheet(output, "results")

        # Test5 has "unknown" as date_result so skip it here
        for pid in ["Test1", "Test2", "Test3", "Test4"]:
            in_date = pd.Timestamp(input_results.loc[input_results["patient_id"] == pid, "date_result"].iloc[0])
            out_date = pd.Timestamp(output_results.loc[output_results["patient_id"] == pid, "date_result"].iloc[0])
            assert (out_date - in_date).days == shift_dict[pid], (
                f"Patient {pid}: results sheet shift differs from linking table"
            )

    def test_shifts_within_specified_range(self, tmp_path):
        _, linking = run_shift(tmp_path, seed=42, min_shift_days=-7, max_shift_days=7)
        shifts = pd.read_csv(linking)["shift_days"]
        assert shifts.between(-7, 7).all()

    def test_placeholder_date_becomes_null_in_output(self, tmp_path):
        """Test5 has "unknown" as date_result — should be null after shifting."""
        output, _ = run_shift(tmp_path, seed=42)
        output_results = read_sheet(output, "results")
        test5_date = output_results.loc[output_results["patient_id"] == "Test5", "date_result"].iloc[0]
        assert pd.isna(test5_date)


class TestReproducibility:
    def test_same_seed_produces_identical_output(self, tmp_path):
        run1 = tmp_path / "run1"
        run2 = tmp_path / "run2"
        run1.mkdir()
        run2.mkdir()
        output1, _ = run_shift(run1, seed=42)
        output2, _ = run_shift(run2, seed=42)

        for sheet in SHEET_CONFIGS:
            pd.testing.assert_frame_equal(
                read_sheet(output1, sheet),
                read_sheet(output2, sheet),
            )

    def test_linking_table_reproduces_same_shifts_on_new_file(self, tmp_path):
        """Saving a linking table and reloading it should produce identical dates."""
        output1, linking = run_shift(tmp_path, seed=42)

        output2 = tmp_path / "output2.xlsx"
        shift_excel_dates(
            str(INPUT_FILE),
            str(output2),
            patient_sheet="patients",
            patient_id_col="patient_id",
            sheet_configs=SHEET_CONFIGS,
            linking_table_path=str(linking),
            linking_table_output=str(tmp_path / "linking2.csv"),
        )

        for sheet in SHEET_CONFIGS:
            pd.testing.assert_frame_equal(
                read_sheet(output1, sheet),
                read_sheet(output2, sheet),
            )
