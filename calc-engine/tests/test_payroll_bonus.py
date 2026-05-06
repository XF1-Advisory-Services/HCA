import unittest

from app.payroll_headcount import calculate_payroll_outputs
from app.schemas import ModelConfig


class PayrollBonusAccrualTests(unittest.TestCase):
    def test_bonus_accrual_by_department(self):
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
                "Start Date": "2026-01-01",
                "Termination Date": "2099-12-31",
                "Bonus Plan": "Executive Plan",
                "Bonus %": 0.10,
                "Bonus $": 0,
                "2026": 120000,
                "2027": 120000,
            },
            {
                "EmployeeID": "E2",
                "__hcaStoreDetail": True,
                "FS_Category": "OpEx",
                "Status": "Domestic",
                "Department": "Sales",
                "Start Date": "2026-01-01",
                "Termination Date": "2026-07-15",
                "Bonus Plan": "MBO Plan - Fixed Bonus",
                "Bonus %": 0,
                "Bonus $": 12000,
                "2026": 100000,
                "2027": 100000,
            },
            {
                "EmployeeID": "E3",
                "__hcaStoreDetail": False,
                "FS_Category": "OpEx",
                "Status": "Domestic",
                "Department": "G&A",
                "Start Date": "2026-01-01",
                "Termination Date": "2099-12-31",
                "Bonus Plan": "Services Plan",
                "Bonus %": 0.20,
                "Bonus $": 0,
                "2026": 90000,
                "2027": 90000,
            },
        ]
        assumptions = {
            "bonus": {
                "cap": 2,
                "netNewArrAchieved": [1.20, 1.20, 1.20],
                "burnMultipleAchieved": [0.80, 0.80, 0.80],
                "executivePlan": {
                    "netNewArrWeight": 0.75,
                    "burnMultipleWeight": 0.25,
                },
                "incentivePlan": {
                    "netNewArrWeight": 1,
                    "burnMultipleWeight": 0,
                },
            }
        }

        outputs = calculate_payroll_outputs(headers, rows, model, assumptions)

        self.assertEqual(
            outputs["bonusAccrual"]["table"],
            [
                ["Department", "Apr 2026", "May 2026", "Jun 2026"],
                ["G&A", 0, 0, 0],
                ["Sales", 2100, 1100, 1100],
            ],
        )
        self.assertEqual(
            outputs["bonusPayout"]["table"],
            [
                ["Department", "Apr 2026", "May 2026", "Jun 2026"],
                ["G&A", 0, 0, 0],
                ["Sales", 0, 2100, 0],
            ],
        )
        self.assertIn(
            {
                "unit_id": "E1",
                "department": "Sales",
                "period_end_date": "2026-05-31",
                "output_key": "payroll.output.bonus_payout",
                "value": 1100,
            },
            outputs["detailRows"],
        )
        self.assertEqual(
            [
                row
                for row in outputs["detailRows"]
                if row["unit_id"] == "E3"
            ],
            [],
        )
        self.assertIn(
            {
                "unit_id": "E2",
                "department": "Sales",
                "period_end_date": "2026-05-31",
                "output_key": "payroll.output.bonus_payout",
                "value": 1000,
            },
            outputs["detailRows"],
        )
        self.assertIn(
            {
                "unit_id": "E1",
                "department": "Sales",
                "period_end_date": "2026-05-31",
                "output_key": "payroll.output.base_salary_domestic",
                "value": 10000,
            },
            outputs["detailRows"],
        )


if __name__ == "__main__":
    unittest.main()
