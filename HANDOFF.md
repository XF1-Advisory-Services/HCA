# Heavy Calc Assist Handoff

## Current State

Heavy Calc Assist is an Excel Desktop Office.js add-in backed by a FastAPI calc engine.

The current MVP reads Payroll input data from the workbook, sends included rows to the backend, calculates payroll outputs, and writes the results back to the workbook.

If `DATABASE_URL` is configured on the backend, each Payroll Recalc also stores selected latest employee/month/output detail rows in Postgres. The task pane `User ID` field supplies the MVP `user_key`.

It does not yet attempt to replace the full payroll model. Current implemented outputs are:

- Headcount / FTE
- Base salary total
- Base salary domestic
- Base salary international
- Base salary COGS
- Medical benefits
- 401k benefits
- Other benefits
- Bonus accrual
- Bonus payout

## Deployment

Frontend add-in:

```text
https://xf1-advisory-services.github.io/HCA/taskpane.html
```

Backend:

```text
https://hca-calc-engine.onrender.com
```

Optional backend environment variable for detail storage:

```text
DATABASE_URL=<Render Postgres internal database URL>
```

Health check:

```text
https://hca-calc-engine.onrender.com/health
```

Microsoft 365 Admin Center manifest:

```text
https://raw.githubusercontent.com/XF1-Advisory-Services/HCA/main/deploy/m365/hca-production-manifest.xml
```

The manifest version is currently:

```text
1.0.3.0
```

If task pane URL, icon URL, display name, permissions, or app metadata changes, bump the manifest version before updating Admin Center.

## Repo Structure

```text
excel-addin/      Office.js task pane add-in
calc-engine/      FastAPI backend
deploy/m365/      Production Microsoft 365 manifest
shared/           Payload notes
README.md         Setup/deploy instructions
HANDOFF.md        Project context for future work
```

## Key Files

Frontend:

```text
excel-addin/src/taskpane/taskpane.js
excel-addin/taskpane.html
excel-addin/src/taskpane/taskpane.css
```

Backend:

```text
calc-engine/app/main.py
calc-engine/app/schemas.py
calc-engine/app/payroll_headcount.py
calc-engine/app/detail_store.py
```

Production manifest:

```text
deploy/m365/hca-production-manifest.xml
```

GitHub Pages workflow:

```text
.github/workflows/deploy-addin.yml
```

Render blueprint:

```text
render.yaml
```

## Workbook Config Contract

The add-in reads the workbook-level named range `HCA.Engine.Config` in one bulk operation. It does not hardcode Config cell addresses.

The named range currently refers to:

```text
Config!$B$2:$G$150
```

The Config range must have these headers:

```text
Section | Type | Key | Description | Value | Value Type
```

The `Key` column is the engine contract. `Section`, `Type`, and `Description` are for readability. Required current keys:

```text
model.last_actuals_date
model.model_end_date
model.financial_year_end_month
payroll.benefits.medical.domestic
payroll.benefits.medical.international
payroll.benefits.401k.domestic
payroll.benefits.401k.international
payroll.benefits.other.domestic
payroll.benefits.other.international
payroll.net_new_ARR_achieved
payroll.burn_multiple_achieved
payroll.bonus_cap
payroll.exec_bonus_NNAR_weight
payroll.exec_bonus_burn_multiple_weight
payroll.incentive_bonus_NNAR_weight
payroll.incentive_bonus_burn_multiple_weight
payroll.filter_column
payroll.store_filter_column
payroll.data_range
payroll.headers_range
payroll.output.headcount
payroll.output.base_salary_total
payroll.output.base_salary_domestic
payroll.output.base_salary_international
payroll.output.base_salary_cogs
payroll.output.medical
payroll.output.401k
payroll.output.other_benefits
payroll.output.bonus_accrual
payroll.output.bonus_payout
```

Example range values:

```text
payroll.filter_column                   PayrollData!R:R
payroll.store_filter_column             PayrollData!S:S
payroll.data_range                      PayrollData!B5:S1531
payroll.headers_range                   PayrollData!B4:S4
payroll.output.headcount                HCA_Output!E4
payroll.output.base_salary_total        HCA_Output!E17
payroll.output.base_salary_domestic     HCA_Output!E30
payroll.output.base_salary_international HCA_Output!E44
payroll.output.base_salary_cogs         HCA_Output!E57
payroll.output.medical                  HCA_Output!E70
payroll.output.401k                     HCA_Output!E83
payroll.output.other_benefits           HCA_Output!E96
payroll.net_new_ARR_achieved            C_Payroll!BP28:CN28
payroll.burn_multiple_achieved          C_Payroll!BP29:CN29
payroll.output.bonus_accrual            HCA_Output!E110
payroll.output.bonus_payout             HCA_Output!E124
```

## PayrollData Contract

Current calculation code assumes this fixed input layout inside the configured Payroll data range:

```text
B  EmployeeID
C  FS_Category
D  Status
E  Department
F  Start Date
G  Termination Date
M:Q Annual base salary assumptions by financial year label
R  Include in LOAD
S  Store Flag
```

In backend zero-based payload field indexes:

```text
FS_Category      1
Status           2
Department       3
Start Date       4
Termination Date 5
```

The add-in filters rows where the configured `payroll.filter_column` equals `1`.

The configured `payroll.store_filter_column` controls employee-level cloud detail storage only. Rows where Store Flag equals `1` are saved to Postgres detail storage. Rows where Store Flag is blank or `0` still contribute to the department-level Excel outputs, but their employee/month detail rows are not saved.

## Calculation Rules

### Timeline

The add-in derives monthly forecast periods from:

```text
calculation start = month after Last actuals date
calculation end   = Model end date
```

Financial year is derived from `Financial year end month`.

If FY end month is `4`, then:

```text
Apr-2026 = FY2026
May-2026 through Apr-2027 = FY2027
```

### Headcount / FTE

For each employee/month:

```text
active_start = max(employee start date, month start date)
active_end   = min(employee end date, month end date)
FTE          = active days / days in month
```

If there is no overlap, FTE is `0`.

Blank termination date means active through model end date.

Output is grouped by Department and rounded to 2 decimals.

### Base Salary

For each employee/month:

```text
monthly base salary = annual salary for period financial year / 12 * FTE
```

Salary field lookup by financial year is precomputed once per period, not repeatedly discovered for every employee row.

Outputs:

- total: all included employees
- domestic: `Status = Domestic`
- international: `Status = International`
- COGS: `FS_Category = COS`, regardless of Status

### Benefits

Benefit assumptions are monthly constants by Status.

For each employee/month:

```text
if FTE > 0:
    benefit = full monthly benefit assumption for employee Status
else:
    benefit = 0
```

Benefits are not prorated by FTE.

Outputs:

- Medical
- 401k
- Other Benefits

Each output is grouped by Department.

### Bonus Accrual

Bonus accrual is calculated by employee/month and grouped by Department.

```text
if Bonus $ > 0:
    monthly bonus base = Bonus $ / 12
else:
    monthly bonus base = annual salary for financial year / 12 * Bonus %
```

Plan multipliers:

- Customer Success Plan: 1.0
- MBO Plan - Fixed Bonus: 1.0
- Executive Plan: performance-based using Executive weights
- Halcyon Incentive Bonus: performance-based using Incentive weights
- na, No Plan, Services Plan, blank, or unknown plan: 0

Performance plans use:

```text
min(Bonus Cap, Net New ARR Achieved %) * Net New ARR Weight
+ min(Bonus Cap, Burn Multiple Achieved %) * Burn Multiple Weight
```

The worked-in-month gate is `1` when FTE is above `0`; bonus accrual is not prorated by FTE. The bonus accrual flag is `1` only when the final eligible bonus cycle end is strictly after the current period end.

### Bonus Payout

Bonus payout is calculated by employee/month and grouped by Department.

Payout months are:

```text
February, May, August, November
```

For each employee in a payout month:

```text
bonus payout = sum of available prior 3 forecast-month bonus accruals
```

The engine does not backfill pre-forecast accruals. For example, if the forecast starts in April 2026, May 2026 payout includes April 2026 bonus accrual only.

## Detail Storage

Detail storage is latest-run-only by user. The backend creates tables automatically when `DATABASE_URL` is present:

```text
calc_runs
calcs_detail_outputs
```

`calcs_detail_outputs` stores one row per employee, forecast month, and output key:

```text
user_key
unit_id
department
period_end_date
output_key
value
```

Output keys are the Config keys, for example:

```text
payroll.output.base_salary_total
payroll.output.bonus_payout
```

On every saved run, the backend deletes the previous run for the same `user_key` and bulk inserts the new detail rows. Zero-value detail rows are filtered out before saving. The workbook Store Flag also limits which employees produce saved detail rows. If `User ID` is blank, `DATABASE_URL` is missing, or the database save fails, Payroll Recalc still returns Excel outputs and reports the detail save status in the response.

The Excel custom function is:

```excel
=HCA.LOAD_DETAIL("payroll.output.base_salary_total", C1, B10)
```

Arguments are output key, any date in the target forecast month, and employee ID. The function normalizes the date to month-end and returns `0` when the selected detail row was not saved.

The response includes timings for calculation, detail save, and total backend time. The task pane logs calculation and detail save timing after each recalc.

## Development Workflow

After changes, run:

```powershell
cd excel-addin
npm run build
npx office-addin-manifest validate ..\deploy\m365\hca-production-manifest.xml
```

Backend syntax check:

```powershell
cd calc-engine
python -m py_compile app\main.py app\schemas.py app\payroll_headcount.py
```

Then commit and push to:

```text
https://github.com/XF1-Advisory-Services/HCA.git
```

GitHub Actions deploys the add-in to GitHub Pages.

Render deploys the backend from the company repo.

## Important Notes

- Do not commit workbook files. `.gitignore` excludes Excel files.
- Do not change the production manifest ID unless intentionally creating a new add-in identity.
- If the Microsoft 365 Admin Center app needs a manifest refresh, update the existing app using the raw manifest URL above.
- If only JavaScript/backend code changes and manifest URLs/metadata do not change, Admin Center usually does not need a manifest update.
- Render free tier may sleep; first backend request can be slow.

## Suggested Next Work

- Add output section titles/formatting in `HCA_Output`.
- Add warnings in task pane for skipped rows or unknown Status values.
- Move fixed PayrollData field indexes to Config if the model layout becomes less stable.
- Add unit tests for payroll calculations.
- Add remaining payroll lines such as severance, payroll taxes, or allocation logic.
