import unittest

from app.payroll_headcount import calculate_payroll_outputs
from app.schemas import ModelConfig


class PayrollSeveranceTests(unittest.TestCase):
    def test_severance_is_paid_once_in_termination_month(self):
        model = ModelConfig(
            lastActualsDate="2026-03-31",
            modelEndDate="2026-06-30",
            calculationStartDate="2026-04-30",
            calculationEndDate="2026-06-30",
            calculationMonths=3,
            financialYearEndMonth=4,
            periods=[
                {"date": "2026-04-30", "label": "Apr 2026", "financialYear": 2026},
                {"date": "2026-05-31", "label": "May 2026", "financialYear": 2027},
                {"date": "2026-06-30", "label": "Jun 2026", "financialYear": 2027},
            ],
        )
        headers = [
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
            "2027",
        ]
        rows = [
            {
                "EmployeeID": "E1",
                "__hcaStoreDetail": True,
                "FS_Category": "OpEx",
                "Status": "Domestic",
                "Department": "Sales",
                "Start Date": "2025-01-01",
                "Termination Date": "2026-05-15",
                "Bonus Plan": "na",
                "Bonus %": 0,
                "Bonus $": 0,
                "Severance Pay": 12000,
                "2026": 120000,
                "2027": 120000,
            },
            {
                "EmployeeID": "E2",
                "__hcaStoreDetail": True,
                "FS_Category": "OpEx",
                "Status": "Domestic",
                "Department": "Sales",
                "Start Date": "2025-01-01",
                "Termination Date": "2026-07-01",
                "Bonus Plan": "na",
                "Bonus %": 0,
                "Bonus $": 0,
                "Severance Pay": 5000,
                "2026": 120000,
                "2027": 120000,
            },
        ]

        outputs = calculate_payroll_outputs(headers, rows, model, assumptions={})

        self.assertEqual(
            outputs["severance"]["table"],
            [
                ["Department", "Apr 2026", "May 2026", "Jun 2026"],
                ["Sales", 0, 12000, 0],
            ],
        )
        self.assertIn(
            {
                "unit_id": "E1",
                "department": "Sales",
                "period_end_date": "2026-05-31",
                "output_key": "payroll.output.severance",
                "value": 12000,
            },
            outputs["detailRows"],
        )


if __name__ == "__main__":
    unittest.main()
