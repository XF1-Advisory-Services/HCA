import unittest
from unittest.mock import patch

from app.detail_store import (
    build_detail_records,
    close_detail_store,
    initialize_detail_store,
    load_detail_value,
    nonzero_detail_rows,
    save_latest_run,
)


class DetailStoreTests(unittest.TestCase):
    def test_build_detail_records_adds_run_and_user_fields(self):
        rows = [
            {
                "unit_id": "E1",
                "department": "Sales",
                "period_end_date": "2026-05-31",
                "output_key": "payroll.output.bonus_payout",
                "value": 1100,
            }
        ]

        records = build_detail_records("run-1", "user@example.com", rows)

        self.assertEqual(
            records,
            [
                (
                    "run-1",
                    "user@example.com",
                    "E1",
                    "Sales",
                    "2026-05-31",
                    "payroll.output.bonus_payout",
                    1100,
                )
            ],
        )

    def test_nonzero_detail_rows_removes_zero_values(self):
        rows = [
            {
                "unit_id": "E1",
                "department": "Sales",
                "period_end_date": "2026-05-31",
                "output_key": "payroll.output.bonus_payout",
                "value": 0,
            },
            {
                "unit_id": "E1",
                "department": "Sales",
                "period_end_date": "2026-05-31",
                "output_key": "payroll.output.base_salary_total",
                "value": 10000,
            },
        ]

        self.assertEqual(
            nonzero_detail_rows(rows),
            [
                {
                    "unit_id": "E1",
                    "department": "Sales",
                    "period_end_date": "2026-05-31",
                    "output_key": "payroll.output.base_salary_total",
                    "value": 10000,
                }
            ],
        )

    def test_load_detail_value_returns_zero_when_database_is_not_configured(self):
        with patch.dict("os.environ", {}, clear=True):
            result = load_detail_value(
                "user@example.com",
                "payroll.output.base_salary_total",
                "2026-05-31",
                "E1",
                database_url=None,
            )

        self.assertEqual(
            result,
            {
                "status": "skipped",
                "reason": "database_not_configured",
                "value": 0,
            },
        )

    def test_initialize_detail_store_skips_without_database_url(self):
        with patch.dict("os.environ", {}, clear=True):
            result = initialize_detail_store(database_url=None)

        self.assertEqual(
            result,
            {"status": "skipped", "reason": "database_not_configured"},
        )

    def test_close_detail_store_without_pool_is_safe(self):
        close_detail_store()
        close_detail_store()

    def test_save_latest_run_skips_without_database_url(self):
        with patch.dict("os.environ", {}, clear=True):
            result = save_latest_run(
                "user@example.com",
                payload=None,
                detail_rows=[],
                database_url=None,
            )

        self.assertEqual(result["status"], "skipped")
        self.assertEqual(result["reason"], "database_not_configured")
        self.assertEqual(result["rowsPrepared"], 0)
        self.assertEqual(result["rowsSaved"], 0)


if __name__ == "__main__":
    unittest.main()
