"""Unit tests for nuh_helper.date_shift"""

from datetime import date, datetime

import pandas as pd
import pytest

from nuh_helper.date_shift import (
    _normalize_patient_id,
    _parse_date_value,
    apply_date_shifts,
    generate_shift_mappings,
    load_shift_mappings,
)


class TestNormalizePatientId:
    def test_none_returns_none(self):
        assert _normalize_patient_id(None) is None

    def test_nan_returns_none(self):
        assert _normalize_patient_id(float("nan")) is None

    def test_empty_string_returns_none(self):
        assert _normalize_patient_id("") is None

    def test_whitespace_only_returns_none(self):
        assert _normalize_patient_id("   ") is None

    def test_strips_whitespace(self):
        assert _normalize_patient_id("  P001  ") == "P001"

    def test_plain_string(self):
        assert _normalize_patient_id("P001") == "P001"

    def test_integer_converts_to_string(self):
        assert _normalize_patient_id(12345) == "12345"


class TestParseDateValue:
    def test_none_returns_none(self):
        assert _parse_date_value(None) is None

    def test_nan_returns_none(self):
        assert _parse_date_value(float("nan")) is None

    def test_empty_string_returns_none(self):
        assert _parse_date_value("") is None

    @pytest.mark.parametrize(
        "placeholder",
        ["unknown", "Unknown", "unk", "unkown", "n/a", "none", "null"],
    )
    def test_placeholder_strings_return_none(self, placeholder):
        assert _parse_date_value(placeholder) is None

    def test_datetime_object(self):
        assert _parse_date_value(datetime(2023, 1, 15)) == pd.Timestamp("2023-01-15")

    def test_date_object(self):
        assert _parse_date_value(date(2023, 1, 15)) == pd.Timestamp("2023-01-15")

    def test_timestamp_object(self):
        ts = pd.Timestamp("2023-01-15")
        assert _parse_date_value(ts) == ts

    def test_yyyy_mm_dd_string(self):
        assert _parse_date_value("2023-01-15") == pd.Timestamp("2023-01-15")

    def test_yyyy_dd_mm_string(self):
        # 15 cannot be a month, so this is unambiguously day-first
        assert _parse_date_value("2023-15-01") == pd.Timestamp("2023-01-15")

    def test_dd_mm_yyyy_string(self):
        assert _parse_date_value("15-01-2023") == pd.Timestamp("2023-01-15")

    def test_mm_dd_yyyy_string(self):
        # 15 cannot be a month, so this is unambiguously day-second
        assert _parse_date_value("01-15-2023") == pd.Timestamp("2023-01-15")


class TestGenerateShiftMappings:
    def test_returns_correct_columns(self):
        result = generate_shift_mappings(["P001"])
        assert list(result.columns) == ["patient_id", "shift_days"]

    def test_one_row_per_patient(self):
        ids = ["P001", "P002", "P003"]
        result = generate_shift_mappings(ids)
        assert len(result) == 3
        assert list(result["patient_id"]) == ids

    def test_reproducible_with_seed(self):
        ids = ["P001", "P002", "P003"]
        result1 = generate_shift_mappings(ids, seed=42)
        result2 = generate_shift_mappings(ids, seed=42)
        assert list(result1["shift_days"]) == list(result2["shift_days"])

    def test_different_seeds_differ(self):
        ids = ["P001", "P002", "P003"]
        result1 = generate_shift_mappings(ids, seed=1)
        result2 = generate_shift_mappings(ids, seed=2)
        assert list(result1["shift_days"]) != list(result2["shift_days"])

    def test_shifts_within_specified_range(self):
        ids = [f"P{i:03d}" for i in range(100)]
        result = generate_shift_mappings(
            ids, min_shift_days=-7, max_shift_days=7, seed=42
        )
        assert result["shift_days"].between(-7, 7).all()

    def test_empty_patient_list(self):
        result = generate_shift_mappings([])
        assert len(result) == 0


class TestLoadShiftMappings:
    def test_loads_valid_csv(self, tmp_path):
        csv_file = tmp_path / "shifts.csv"
        csv_file.write_text("patient_id,shift_days\nP001,5\nP002,-3\n")
        result = load_shift_mappings(str(csv_file))
        assert list(result["patient_id"]) == ["P001", "P002"]
        assert list(result["shift_days"]) == [5, -3]

    def test_raises_on_missing_columns(self, tmp_path):
        csv_file = tmp_path / "bad.csv"
        csv_file.write_text("id,days\nP001,5\n")
        with pytest.raises(ValueError, match="patient_id.*shift_days"):
            load_shift_mappings(str(csv_file))


class TestApplyDateShifts:
    def _make_mappings(self, patient_id: str, shift_days: int) -> pd.DataFrame:
        return pd.DataFrame({"patient_id": [patient_id], "shift_days": [shift_days]})

    def test_shifts_date_forward(self):
        df = pd.DataFrame({"patient_id": ["P001"], "visit_date": ["2023-01-15"]})
        result = apply_date_shifts(
            df, "patient_id", ["visit_date"], self._make_mappings("P001", 10)
        )
        assert result["visit_date"].iloc[0] == date(2023, 1, 25)

    def test_shifts_date_backward(self):
        df = pd.DataFrame({"patient_id": ["P001"], "visit_date": ["2023-06-01"]})
        result = apply_date_shifts(
            df, "patient_id", ["visit_date"], self._make_mappings("P001", -5)
        )
        assert result["visit_date"].iloc[0] == date(2023, 5, 27)

    def test_zero_shift_leaves_date_unchanged(self):
        df = pd.DataFrame({"patient_id": ["P001"], "visit_date": ["2023-01-15"]})
        result = apply_date_shifts(
            df, "patient_id", ["visit_date"], self._make_mappings("P001", 0)
        )
        assert result["visit_date"].iloc[0] == date(2023, 1, 15)

    def test_unknown_patient_leaves_date_unchanged(self):
        df = pd.DataFrame({"patient_id": ["P999"], "visit_date": ["2023-01-15"]})
        result = apply_date_shifts(
            df, "patient_id", ["visit_date"], self._make_mappings("P001", 10)
        )
        assert result["visit_date"].iloc[0] == date(2023, 1, 15)

    def test_placeholder_date_becomes_none(self):
        df = pd.DataFrame({"patient_id": ["P001"], "visit_date": ["Unknown"]})
        result = apply_date_shifts(
            df, "patient_id", ["visit_date"], self._make_mappings("P001", 10)
        )
        assert result["visit_date"].iloc[0] is None

    def test_none_date_stays_none(self):
        df = pd.DataFrame({"patient_id": ["P001"], "visit_date": [None]})
        result = apply_date_shifts(
            df, "patient_id", ["visit_date"], self._make_mappings("P001", 10)
        )
        assert result["visit_date"].iloc[0] is None

    def test_missing_date_column_is_skipped(self):
        df = pd.DataFrame({"patient_id": ["P001"], "visit_date": ["2023-01-15"]})
        mappings = self._make_mappings("P001", 10)
        # "nonexistent" should be silently skipped; "visit_date" still shifted
        result = apply_date_shifts(
            df, "patient_id", ["nonexistent", "visit_date"], mappings
        )
        assert result["visit_date"].iloc[0] == date(2023, 1, 25)

    def test_strips_whitespace_from_patient_ids(self):
        df = pd.DataFrame({"patient_id": ["  P001  "], "visit_date": ["2023-01-15"]})
        result = apply_date_shifts(
            df, "patient_id", ["visit_date"], self._make_mappings("P001", 10)
        )
        assert result["visit_date"].iloc[0] == date(2023, 1, 25)

    def test_multiple_patients_shifted_independently(self):
        df = pd.DataFrame(
            {
                "patient_id": ["P001", "P002"],
                "visit_date": ["2023-01-01", "2023-01-01"],
            }
        )
        mappings = pd.DataFrame({"patient_id": ["P001", "P002"], "shift_days": [5, -5]})
        result = apply_date_shifts(df, "patient_id", ["visit_date"], mappings)
        assert result["visit_date"].iloc[0] == date(2023, 1, 6)
        assert result["visit_date"].iloc[1] == date(2022, 12, 27)

    def test_result_dates_are_date_not_datetime(self):
        df = pd.DataFrame({"patient_id": ["P001"], "visit_date": ["2023-01-15"]})
        result = apply_date_shifts(
            df, "patient_id", ["visit_date"], self._make_mappings("P001", 0)
        )
        val = result["visit_date"].iloc[0]
        assert isinstance(val, date)
        assert not isinstance(val, datetime)

    def test_does_not_mutate_input(self):
        df = pd.DataFrame({"patient_id": ["P001"], "visit_date": ["2023-01-15"]})
        original_visit = df["visit_date"].iloc[0]
        apply_date_shifts(
            df, "patient_id", ["visit_date"], self._make_mappings("P001", 10)
        )
        assert df["visit_date"].iloc[0] == original_visit
