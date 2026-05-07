from calendar import monthrange
from collections import defaultdict
from datetime import date, datetime, timedelta
from typing import Any

from app.schemas import ModelConfig


DEPARTMENT_FIELD_INDEX = 3
EMPLOYEE_ID_FIELD_INDEX = 0
START_DATE_FIELD_INDEX = 4
END_DATE_FIELD_INDEX = 5
FS_CATEGORY_FIELD_INDEX = 1
STATUS_FIELD_INDEX = 2
BONUS_PLAN_FIELD_INDEX = 6
BONUS_PERCENT_FIELD_INDEX = 7
BONUS_FIXED_FIELD_INDEX = 8
SEVERANCE_PAY_FIELD_INDEX = 10
FAR_FUTURE_DATE = date(2099, 12, 31)
BONUS_PAYOUT_MONTHS = {2, 5, 8, 11}
DETAIL_OUTPUT_KEYS = {
    "headcount": "payroll.output.headcount",
    "base_salary_total": "payroll.output.base_salary_total",
    "base_salary_domestic": "payroll.output.base_salary_domestic",
    "base_salary_international": "payroll.output.base_salary_international",
    "base_salary_cogs": "payroll.output.base_salary_cogs",
    "medical": "payroll.output.medical",
    "retirement_401k": "payroll.output.401k",
    "other_benefits": "payroll.output.other_benefits",
    "bonus_accrual": "payroll.output.bonus_accrual",
    "bonus_payout": "payroll.output.bonus_payout",
    "severance": "payroll.output.severance",
}
STORE_DETAIL_FIELD = "__hcaStoreDetail"


def calculate_payroll_outputs(
    headers: list[str],
    rows: list[dict[str, Any]],
    model: ModelConfig,
    assumptions: dict[str, Any],
) -> dict[str, Any]:
    employee_id_field = field_name(headers, EMPLOYEE_ID_FIELD_INDEX, "employee id")
    department_field = field_name(headers, DEPARTMENT_FIELD_INDEX, "department")
    start_date_field = field_name(headers, START_DATE_FIELD_INDEX, "start date")
    end_date_field = field_name(headers, END_DATE_FIELD_INDEX, "end date")
    fs_category_field = field_name(headers, FS_CATEGORY_FIELD_INDEX, "FS category")
    status_field = field_name(headers, STATUS_FIELD_INDEX, "status")
    bonus_plan_field = field_name(headers, BONUS_PLAN_FIELD_INDEX, "bonus plan")
    bonus_percent_field = field_name(headers, BONUS_PERCENT_FIELD_INDEX, "bonus percent")
    bonus_fixed_field = field_name(headers, BONUS_FIXED_FIELD_INDEX, "bonus fixed amount")
    severance_pay_field = field_name(
        headers, SEVERANCE_PAY_FIELD_INDEX, "severance pay"
    )
    periods = [
        {
            "date": parse_iso_date(period.date),
            "label": period.label,
            "financialYear": period.financialYear,
        }
        for period in model.periods
    ]
    salary_fields = salary_field_by_period(headers, periods)
    headcount_totals: dict[str, list[float]] = defaultdict(
        lambda: [0.0] * len(periods)
    )
    base_salary_totals: dict[str, list[float]] = defaultdict(
        lambda: [0.0] * len(periods)
    )
    domestic_salary_totals: dict[str, list[float]] = defaultdict(
        lambda: [0.0] * len(periods)
    )
    international_salary_totals: dict[str, list[float]] = defaultdict(
        lambda: [0.0] * len(periods)
    )
    cogs_salary_totals: dict[str, list[float]] = defaultdict(
        lambda: [0.0] * len(periods)
    )
    benefit_rates = parse_benefit_rates(assumptions)
    medical_totals: dict[str, list[float]] = defaultdict(lambda: [0.0] * len(periods))
    retirement_401k_totals: dict[str, list[float]] = defaultdict(
        lambda: [0.0] * len(periods)
    )
    other_benefits_totals: dict[str, list[float]] = defaultdict(
        lambda: [0.0] * len(periods)
    )
    bonus_assumptions = parse_bonus_assumptions(assumptions)
    bonus_accrual_totals: dict[str, list[float]] = defaultdict(
        lambda: [0.0] * len(periods)
    )
    bonus_payout_totals: dict[str, list[float]] = defaultdict(
        lambda: [0.0] * len(periods)
    )
    severance_totals: dict[str, list[float]] = defaultdict(
        lambda: [0.0] * len(periods)
    )
    detail_rows: list[dict[str, Any]] = []
    skipped_rows = 0

    for row in rows:
        unit_id = normalize_text(row.get(employee_id_field))
        department = normalize_text(row.get(department_field))
        start_date = parse_date_value(row.get(start_date_field))
        raw_end_date = parse_date_value(row.get(end_date_field))
        end_date = raw_end_date or parse_iso_date(model.calculationEndDate)
        bonus_end_date = raw_end_date or FAR_FUTURE_DATE

        if not unit_id or not department or start_date is None:
            skipped_rows += 1
            continue

        status = normalize_key(row.get(status_field))
        fs_category = normalize_key(row.get(fs_category_field))
        bonus_plan = normalize_key(row.get(bonus_plan_field))
        bonus_percent = parse_number(row.get(bonus_percent_field))
        bonus_fixed_amount = parse_number(row.get(bonus_fixed_field))
        severance_pay = parse_number(row.get(severance_pay_field))
        final_bonus_cycle_end = final_eligible_bonus_cycle_end(bonus_end_date)
        employee_bonus_accruals = [0.0] * len(periods)
        store_detail = is_enabled(row.get(STORE_DETAIL_FIELD))

        for index, period in enumerate(periods):
            month_end = period["date"]
            month_start = date(month_end.year, month_end.month, 1)
            fte = monthly_fte(start_date, end_date, month_start, month_end)
            annual_salary = parse_number(row.get(salary_fields[index]))
            monthly_salary = (annual_salary / 12) * fte
            domestic_salary = monthly_salary if status == "domestic" else 0.0
            international_salary = (
                monthly_salary if status == "international" else 0.0
            )
            cogs_salary = monthly_salary if fs_category == "cos" else 0.0
            headcount_totals[department][index] += fte
            base_salary_totals[department][index] += monthly_salary

            if status == "domestic":
                domestic_salary_totals[department][index] += domestic_salary
            elif status == "international":
                international_salary_totals[department][index] += international_salary

            if fs_category == "cos":
                cogs_salary_totals[department][index] += cogs_salary

            status_rates = benefit_rates.get(status, ZERO_BENEFIT_RATES)
            benefit_multiplier = 1.0 if fte > 0 else 0.0
            medical = benefit_multiplier * status_rates["medical"]
            retirement_401k = benefit_multiplier * status_rates["retirement401k"]
            other_benefits = benefit_multiplier * status_rates["otherBenefits"]
            medical_totals[department][index] += medical
            retirement_401k_totals[department][index] += retirement_401k
            other_benefits_totals[department][index] += other_benefits

            monthly_bonus_base = monthly_bonus_amount(
                monthly_salary,
                bonus_percent,
                bonus_fixed_amount,
            )
            employee_bonus_accrual = (
                monthly_bonus_base
                * bonus_plan_multiplier(bonus_plan, bonus_assumptions, index)
                * benefit_multiplier
                * bonus_accrual_flag(final_bonus_cycle_end, month_end)
            )
            employee_bonus_accruals[index] = employee_bonus_accrual
            bonus_accrual_totals[department][index] += employee_bonus_accrual
            severance = (
                severance_pay
                if is_within_month(raw_end_date, month_start, month_end)
                else 0.0
            )
            severance_totals[department][index] += severance
            period_date = month_end.isoformat()
            if store_detail:
                append_detail_row(
                    detail_rows,
                    unit_id,
                    department,
                    period_date,
                    DETAIL_OUTPUT_KEYS["headcount"],
                    fte,
                )
                append_detail_row(
                    detail_rows,
                    unit_id,
                    department,
                    period_date,
                    DETAIL_OUTPUT_KEYS["base_salary_total"],
                    monthly_salary,
                )
                append_detail_row(
                    detail_rows,
                    unit_id,
                    department,
                    period_date,
                    DETAIL_OUTPUT_KEYS["base_salary_domestic"],
                    domestic_salary,
                )
                append_detail_row(
                    detail_rows,
                    unit_id,
                    department,
                    period_date,
                    DETAIL_OUTPUT_KEYS["base_salary_international"],
                    international_salary,
                )
                append_detail_row(
                    detail_rows,
                    unit_id,
                    department,
                    period_date,
                    DETAIL_OUTPUT_KEYS["base_salary_cogs"],
                    cogs_salary,
                )
                append_detail_row(
                    detail_rows,
                    unit_id,
                    department,
                    period_date,
                    DETAIL_OUTPUT_KEYS["medical"],
                    medical,
                )
                append_detail_row(
                    detail_rows,
                    unit_id,
                    department,
                    period_date,
                    DETAIL_OUTPUT_KEYS["retirement_401k"],
                    retirement_401k,
                )
                append_detail_row(
                    detail_rows,
                    unit_id,
                    department,
                    period_date,
                    DETAIL_OUTPUT_KEYS["other_benefits"],
                    other_benefits,
                )
                append_detail_row(
                    detail_rows,
                    unit_id,
                    department,
                    period_date,
                    DETAIL_OUTPUT_KEYS["bonus_accrual"],
                    employee_bonus_accrual,
                )
                append_detail_row(
                    detail_rows,
                    unit_id,
                    department,
                    period_date,
                    DETAIL_OUTPUT_KEYS["severance"],
                    severance,
                )

        for index, period in enumerate(periods):
            if is_bonus_payout_month(period["date"]):
                employee_bonus_payout = sum(
                    employee_bonus_accruals[max(0, index - 3) : index]
                )
                bonus_payout_totals[department][index] += employee_bonus_payout
            else:
                employee_bonus_payout = 0.0

            if store_detail:
                append_detail_row(
                    detail_rows,
                    unit_id,
                    department,
                    period["date"].isoformat(),
                    DETAIL_OUTPUT_KEYS["bonus_payout"],
                    employee_bonus_payout,
                )

    departments = sorted(headcount_totals)
    domestic_departments = sorted(domestic_salary_totals)
    international_departments = sorted(international_salary_totals)
    cogs_departments = sorted(cogs_salary_totals)
    medical_departments = sorted(medical_totals)
    retirement_401k_departments = sorted(retirement_401k_totals)
    other_benefits_departments = sorted(other_benefits_totals)
    bonus_accrual_departments = sorted(bonus_accrual_totals)
    bonus_payout_departments = sorted(bonus_payout_totals)
    severance_departments = sorted(severance_totals)

    return {
        "headcount": {
            "table": output_table(departments, periods, headcount_totals, decimals=2),
            "departments": departments,
            "periods": serialize_periods(periods),
            "skippedRows": skipped_rows,
            "fieldMap": {
                "department": department_field,
                "startDate": start_date_field,
                "endDate": end_date_field,
            },
        },
        "baseSalary": {
            "total": {
                "table": output_table(
                    departments, periods, base_salary_totals, decimals=0
                ),
                "departments": departments,
            },
            "domestic": {
                "table": output_table(
                    domestic_departments,
                    periods,
                    domestic_salary_totals,
                    decimals=0,
                ),
                "departments": domestic_departments,
            },
            "international": {
                "table": output_table(
                    international_departments,
                    periods,
                    international_salary_totals,
                    decimals=0,
                ),
                "departments": international_departments,
            },
            "cogs": {
                "table": output_table(
                    cogs_departments, periods, cogs_salary_totals, decimals=0
                ),
                "departments": cogs_departments,
            },
            "periods": serialize_periods(periods),
            "salaryFieldByPeriod": salary_fields,
        },
        "benefits": {
            "medical": {
                "table": output_table(
                    medical_departments, periods, medical_totals, decimals=0
                ),
                "departments": medical_departments,
            },
            "retirement401k": {
                "table": output_table(
                    retirement_401k_departments,
                    periods,
                    retirement_401k_totals,
                    decimals=0,
                ),
                "departments": retirement_401k_departments,
            },
            "otherBenefits": {
                "table": output_table(
                    other_benefits_departments,
                    periods,
                    other_benefits_totals,
                    decimals=0,
                ),
                "departments": other_benefits_departments,
            },
            "periods": serialize_periods(periods),
        },
        "bonusAccrual": {
            "table": output_table(
                bonus_accrual_departments,
                periods,
                bonus_accrual_totals,
                decimals=0,
            ),
            "departments": bonus_accrual_departments,
            "periods": serialize_periods(periods),
        },
        "bonusPayout": {
            "table": output_table(
                bonus_payout_departments,
                periods,
                bonus_payout_totals,
                decimals=0,
            ),
            "departments": bonus_payout_departments,
            "periods": serialize_periods(periods),
        },
        "severance": {
            "table": output_table(
                severance_departments,
                periods,
                severance_totals,
                decimals=0,
            ),
            "departments": severance_departments,
            "periods": serialize_periods(periods),
        },
        "detailRows": detail_rows,
    }


def append_detail_row(
    detail_rows: list[dict[str, Any]],
    unit_id: str,
    department: str,
    period_end_date: str,
    output_key: str,
    value: float,
) -> None:
    detail_rows.append(
        {
            "unit_id": unit_id,
            "department": department,
            "period_end_date": period_end_date,
            "output_key": output_key,
            "value": round(value, 6),
        }
    )


def monthly_fte(
    employee_start: date,
    employee_end: date,
    month_start: date,
    month_end: date,
) -> float:
    active_start = max(employee_start, month_start)
    active_end = min(employee_end, month_end)

    if active_start > active_end:
        return 0.0

    active_days = (active_end - active_start).days + 1
    days_in_month = monthrange(month_end.year, month_end.month)[1]
    return active_days / days_in_month


def is_within_month(value: date | None, month_start: date, month_end: date) -> bool:
    return value is not None and month_start <= value <= month_end


def salary_field_by_period(
    headers: list[str],
    periods: list[dict[str, Any]],
) -> list[str | None]:
    available_salary_fields = {str(header).strip(): header for header in headers}
    return [
        available_salary_fields.get(str(period["financialYear"]))
        for period in periods
    ]


def output_table(
    departments: list[str],
    periods: list[dict[str, Any]],
    totals: dict[str, list[float]],
    decimals: int,
) -> list[list[Any]]:
    table = [["Department", *[period["label"] for period in periods]]]
    for department in departments:
        table.append(
            [
                department,
                *[round(value, decimals) for value in totals[department]],
            ]
        )
    return table


def serialize_periods(periods: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        {
            "date": period["date"].isoformat(),
            "label": period["label"],
            "financialYear": period["financialYear"],
        }
        for period in periods
    ]


def field_name(headers: list[str], index: int, fallback: str) -> str:
    if index >= len(headers):
        raise ValueError(f"Payroll data is missing expected field: {fallback}")
    return headers[index]


def normalize_text(value: Any) -> str:
    return str(value or "").strip()


def normalize_key(value: Any) -> str:
    return normalize_text(value).lower()


def is_enabled(value: Any) -> bool:
    if isinstance(value, bool):
        return value

    if isinstance(value, (int, float)):
        return value == 1

    return str(value or "").strip().lower() in ("1", "true", "yes")


def parse_number(value: Any) -> float:
    if value in (None, "", "-"):
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace(",", "")
    if not text or text == "-":
        return 0.0

    if text.endswith("%"):
        return float(text[:-1]) / 100

    return float(text)


ZERO_BENEFIT_RATES = {
    "medical": 0.0,
    "retirement401k": 0.0,
    "otherBenefits": 0.0,
}


def parse_benefit_rates(assumptions: dict[str, Any]) -> dict[str, dict[str, float]]:
    benefits = assumptions.get("benefits", {})

    return {
        "domestic": {
            "medical": parse_number(benefits.get("medical", {}).get("domestic")),
            "retirement401k": parse_number(
                benefits.get("retirement401k", {}).get("domestic")
            ),
            "otherBenefits": parse_number(
                benefits.get("otherBenefits", {}).get("domestic")
            ),
        },
        "international": {
            "medical": parse_number(benefits.get("medical", {}).get("international")),
            "retirement401k": parse_number(
                benefits.get("retirement401k", {}).get("international")
            ),
            "otherBenefits": parse_number(
                benefits.get("otherBenefits", {}).get("international")
            ),
        },
    }


def monthly_bonus_amount(
    monthly_salary: float,
    bonus_percent: float,
    bonus_fixed_amount: float,
) -> float:
    if bonus_fixed_amount > 0:
        return bonus_fixed_amount / 12

    return monthly_salary * bonus_percent


def bonus_plan_multiplier(
    plan: str,
    assumptions: dict[str, Any],
    period_index: int,
) -> float:
    if plan in ("customer success plan", "mbo plan - fixed bonus"):
        return 1.0

    if plan == "executive plan":
        return performance_bonus_multiplier(
            assumptions["cap"],
            assumptions["netNewArrAchieved"],
            assumptions["burnMultipleAchieved"],
            assumptions["executivePlan"],
            period_index,
        )

    if plan == "halcyon incentive bonus":
        return performance_bonus_multiplier(
            assumptions["cap"],
            assumptions["netNewArrAchieved"],
            assumptions["burnMultipleAchieved"],
            assumptions["incentivePlan"],
            period_index,
        )

    return 0.0


def performance_bonus_multiplier(
    cap: float,
    net_new_arr_achieved: list[float],
    burn_multiple_achieved: list[float],
    weights: dict[str, float],
    period_index: int,
) -> float:
    return (
        min(cap, value_at(net_new_arr_achieved, period_index))
        * weights["netNewArrWeight"]
        + min(cap, value_at(burn_multiple_achieved, period_index))
        * weights["burnMultipleWeight"]
    )


def bonus_accrual_flag(final_bonus_cycle_end: date, month_end: date) -> float:
    return 1.0 if final_bonus_cycle_end > month_end else 0.0


def is_bonus_payout_month(month_end: date) -> bool:
    return month_end.month in BONUS_PAYOUT_MONTHS


def final_eligible_bonus_cycle_end(termination_date: date) -> date:
    months_back = (termination_date.month - 2) % 3
    cycle_month = termination_date.month - months_back
    cycle_year = termination_date.year

    if cycle_month <= 0:
        cycle_month += 12
        cycle_year -= 1

    return date(cycle_year, cycle_month, monthrange(cycle_year, cycle_month)[1])


def parse_bonus_assumptions(assumptions: dict[str, Any]) -> dict[str, Any]:
    bonus = assumptions.get("bonus", {})

    return {
        "cap": parse_number(bonus.get("cap")),
        "netNewArrAchieved": parse_number_series(bonus.get("netNewArrAchieved")),
        "burnMultipleAchieved": parse_number_series(
            bonus.get("burnMultipleAchieved")
        ),
        "executivePlan": {
            "netNewArrWeight": parse_number(
                bonus.get("executivePlan", {}).get("netNewArrWeight")
            ),
            "burnMultipleWeight": parse_number(
                bonus.get("executivePlan", {}).get("burnMultipleWeight")
            ),
        },
        "incentivePlan": {
            "netNewArrWeight": parse_number(
                bonus.get("incentivePlan", {}).get("netNewArrWeight")
            ),
            "burnMultipleWeight": parse_number(
                bonus.get("incentivePlan", {}).get("burnMultipleWeight")
            ),
        },
    }


def parse_number_series(values: Any) -> list[float]:
    if not isinstance(values, list):
        return []

    return [parse_number(value) for value in values]


def value_at(values: list[float], index: int) -> float:
    if index >= len(values):
        return 0.0

    return values[index]


def parse_date_value(value: Any) -> date | None:
    if value in (None, ""):
        return None

    if isinstance(value, (int, float)):
        return date(1899, 12, 30) + timedelta(days=int(value))

    if isinstance(value, date):
        return value

    text = str(value).strip()
    if not text:
        return None

    for fmt in ("%Y-%m-%d", "%d-%b-%y", "%d-%b-%Y", "%d/%b/%y", "%d/%b/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass

    raise ValueError(f"Invalid date value: {text}")


def parse_iso_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()
