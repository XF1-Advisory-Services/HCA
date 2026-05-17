import os
import unittest

try:
    from app.main import (
        debug_client_log,
        parse_load_detail_batch_transport,
        payroll_load_detail_batch,
        payroll_load_detail,
        payroll_load_detail_get,
        payroll_load_preview,
    )
    from app.schemas import (
        ClientLogRequest,
        PayrollLoadDetailBatchRequest,
        PayrollLoadDetailRequest,
        PayrollLoadPreviewRequest,
    )
except ModuleNotFoundError:
    debug_client_log = None
    parse_load_detail_batch_transport = None
    payroll_load_detail_batch = None
    payroll_load_detail = None
    payroll_load_detail_get = None
    payroll_load_preview = None
    ClientLogRequest = None
    PayrollLoadDetailBatchRequest = None
    PayrollLoadDetailRequest = None
    PayrollLoadPreviewRequest = None


class PayrollApiTests(unittest.TestCase):
    def test_load_preview_skips_detail_storage_when_database_is_not_configured(self):
        if payroll_load_preview is None:
            self.skipTest("FastAPI test dependency is not installed.")

        previous_database_url = os.environ.pop("DATABASE_URL", None)
        self.addCleanup(self._restore_database_url, previous_database_url)

        response = payroll_load_preview(
            PayrollLoadPreviewRequest.model_validate(
                {
                    "section": "Payroll",
                    "userKey": "user@example.com",
                    "model": {
                        "lastActualsDate": "2026-03-31",
                        "modelEndDate": "2026-04-30",
                        "calculationStartDate": "2026-04-30",
                        "calculationEndDate": "2026-04-30",
                        "calculationMonths": 1,
                        "financialYearEndMonth": 4,
                        "periods": [
                            {
                                "date": "2026-04-30",
                                "label": "Apr 2026",
                                "financialYear": 2026,
                            }
                        ],
                    },
                    "source": {
                        "sheet": "PayrollData",
                        "headerRange": "B4:R4",
                        "dataRange": "B5:R1531",
                        "filterColumn": "R",
                    },
                    "assumptions": {},
                    "metrics": {"totalRows": 1, "includedRows": 1, "loadTimeMs": 1},
                    "headers": [
                        "EmployeeID",
                        "FS_Category",
                        "Status",
                        "Department",
                        "Start Date",
                        "Termination Date",
                        "Bonus Plan",
                        "Bonus %",
                        "Bonus $",
                        "Payroll Case",
                        "Severance Pay",
                        "2026",
                    ],
                    "rows": [
                        {
                            "EmployeeID": "E1",
                            "FS_Category": "OpEx",
                            "Status": "Domestic",
                            "Department": "Sales",
                            "Start Date": "2026-01-01",
                            "Termination Date": "2099-12-31",
                            "Bonus Plan": "na",
                            "Bonus %": 0,
                            "Bonus $": 0,
                            "2026": 120000,
                        }
                    ],
                }
            )
        )

        self.assertEqual(response["detailSave"]["status"], "skipped")
        self.assertEqual(response["detailSave"]["reason"], "database_not_configured")
        self.assertGreaterEqual(response["timings"]["calculationMs"], 0)
        self.assertGreaterEqual(response["timings"]["detailSaveMs"], 0)
        self.assertGreaterEqual(response["timings"]["totalBackendMs"], 0)
        self.assertNotIn("detailRows", response["outputs"])

    def test_load_detail_returns_zero_when_database_is_not_configured(self):
        if payroll_load_detail is None:
            self.skipTest("FastAPI test dependency is not installed.")

        previous_database_url = os.environ.pop("DATABASE_URL", None)
        self.addCleanup(self._restore_database_url, previous_database_url)

        response = payroll_load_detail(
            PayrollLoadDetailRequest.model_validate(
                {
                    "userKey": "user@example.com",
                    "outputKey": "payroll.output.base_salary_total",
                    "periodEndDate": "2026-05-31",
                    "unitId": "E1",
                }
            )
        )

        self.assertEqual(response["status"], "skipped")
        self.assertEqual(response["reason"], "database_not_configured")
        self.assertEqual(response["value"], 0)

    def test_load_detail_get_returns_zero_when_database_is_not_configured(self):
        if payroll_load_detail_get is None:
            self.skipTest("FastAPI test dependency is not installed.")

        previous_database_url = os.environ.pop("DATABASE_URL", None)
        self.addCleanup(self._restore_database_url, previous_database_url)

        response = payroll_load_detail_get(
            userKey="user@example.com",
            outputKey="payroll.output.base_salary_total",
            periodEndDate="2026-05-31",
            unitId="E1",
        )

        self.assertEqual(response["status"], "skipped")
        self.assertEqual(response["reason"], "database_not_configured")
        self.assertEqual(response["value"], 0)

    def test_load_detail_batch_returns_zeros_when_database_is_not_configured(self):
        if payroll_load_detail_batch is None:
            self.skipTest("FastAPI test dependency is not installed.")

        previous_database_url = os.environ.pop("DATABASE_URL", None)
        self.addCleanup(self._restore_database_url, previous_database_url)

        response = payroll_load_detail_batch(
            PayrollLoadDetailBatchRequest.model_validate(
                {
                    "userKey": "user@example.com",
                    "items": [
                        {
                            "outputKey": "payroll.output.base_salary_total",
                            "periodEndDate": "2026-05-31",
                            "unitId": "E1",
                        },
                        {
                            "outputKey": "payroll.output.401k",
                            "periodEndDate": "2026-05-31",
                            "unitId": "E1",
                        },
                    ],
                }
            )
        )

        self.assertEqual(response["status"], "skipped")
        self.assertEqual(response["reason"], "database_not_configured")
        self.assertEqual(response["values"], [0.0, 0.0])
        self.assertEqual(response["foundCount"], 0)
        self.assertEqual(response["timings"]["itemCount"], 2)
        self.assertGreaterEqual(response["timings"]["totalBackendMs"], 0)

    def test_parse_load_detail_batch_transport_accepts_text_payload(self):
        if parse_load_detail_batch_transport is None:
            self.skipTest("FastAPI test dependency is not installed.")

        request = parse_load_detail_batch_transport(
            "text/plain",
            '{"userKey":"user@example.com","items":[{"outputKey":"payroll.output.base_salary_total","periodEndDate":"2026-05-31","unitId":"E1"}]}',
        )

        self.assertEqual(request.userKey, "user@example.com")
        self.assertEqual(len(request.items), 1)
        self.assertEqual(request.items[0].outputKey, "payroll.output.base_salary_total")

    def test_parse_load_detail_batch_transport_accepts_form_payload(self):
        if parse_load_detail_batch_transport is None:
            self.skipTest("FastAPI test dependency is not installed.")

        request = parse_load_detail_batch_transport(
            "application/x-www-form-urlencoded",
            "payload=%7B%22userKey%22%3A%22user%40example.com%22%2C%22items%22%3A%5B%7B%22outputKey%22%3A%22payroll.output.401k%22%2C%22periodEndDate%22%3A%222026-05-31%22%2C%22unitId%22%3A%22E1%22%7D%5D%7D",
        )

        self.assertEqual(request.userKey, "user@example.com")
        self.assertEqual(request.items[0].outputKey, "payroll.output.401k")

    def test_load_detail_batch_rejects_more_than_1000_items(self):
        if PayrollLoadDetailBatchRequest is None:
            self.skipTest("FastAPI test dependency is not installed.")

        with self.assertRaises(ValueError):
            PayrollLoadDetailBatchRequest.model_validate(
                {
                    "userKey": "user@example.com",
                    "items": [
                        {
                            "outputKey": "payroll.output.base_salary_total",
                            "periodEndDate": "2026-05-31",
                            "unitId": f"E{index}",
                        }
                        for index in range(1001)
                    ],
                }
            )

    def test_debug_client_log_records_client_error(self):
        if debug_client_log is None:
            self.skipTest("FastAPI test dependency is not installed.")

        with self.assertLogs("hca.client", level="WARNING") as captured:
            response = debug_client_log(
                ClientLogRequest.model_validate(
                    {
                        "source": "HCA.LOAD_DETAIL",
                        "stage": "backend-fetch",
                        "level": "error",
                        "message": "Failed to fetch",
                        "context": {
                            "outputKey": "payroll.output.401k",
                            "periodEndDate": "2026-04-30",
                            "unitId": "EX18",
                        },
                    }
                )
            )

        self.assertEqual(response, {"status": "logged"})
        self.assertIn("stage=backend-fetch", captured.output[0])

    @staticmethod
    def _restore_database_url(value):
        if value is not None:
            os.environ["DATABASE_URL"] = value


if __name__ == "__main__":
    unittest.main()
