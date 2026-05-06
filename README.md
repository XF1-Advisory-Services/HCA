# External Calc Engine v01

Windows Desktop Excel Office.js add-in plus a Python FastAPI calc engine.

The current version bulk-loads Payroll input rows from Excel, filters rows where `Include in LOAD` equals `1`, sends them to the backend, and writes payroll outputs back to the workbook.

If Postgres is configured, the backend also stores selected latest employee/month/output detail rows for future `HCA.LOAD_DETAIL` lookups.

## Project Structure

- `excel-addin/` - Office.js task pane add-in.
- `calc-engine/` - FastAPI backend.
- `shared/` - Shared payload notes.

## Local Backend

From `calc-engine/`:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m uvicorn app.main:app --host localhost --port 8000 --ssl-keyfile "$env:USERPROFILE\.office-addin-dev-certs\localhost.key" --ssl-certfile "$env:USERPROFILE\.office-addin-dev-certs\localhost.crt"
```

Check:

```powershell
Invoke-RestMethod https://localhost:8000/health
```

Hosted backend:

```text
https://hca-calc-engine.onrender.com
```

## Local Excel Add-in

From `excel-addin/`:

```powershell
npm install
npm run certs
npm run dev
```

The local task pane URL is:

```text
https://localhost:3000/taskpane.html
```

The local manifest is:

```text
excel-addin/manifest.xml
```

## Excel Sideloading

For Windows Desktop Excel, the simple local sideload path is:

1. Create a trusted catalog folder, for example `C:\OfficeAddins\ExternalCalcEngine`.
2. Copy `excel-addin\manifest.xml` into that folder.
3. In Excel, go to `File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs`.
4. Add the folder path and tick `Show in Menu`.
5. Restart Excel.
6. Open the workbook, then use `Insert > My Add-ins > Shared Folder` and select the add-in.

## Workbook Assumptions

The workbook must contain:

- A `Config` sheet.
- A workbook-level named range `HCA.Engine.Config` that points to the Config table.
- Config columns named:
  - `Section`
  - `Type`
  - `Key`
  - `Description`
  - `Value`
  - `Value Type`

The task pane has a `User ID` field. For the MVP, each user should enter their work email address. This is used only to separate latest-run detail storage by user.

Required model keys:

- `model.last_actuals_date`
- `model.model_end_date`
- `model.financial_year_end_month`

Required payroll keys:

- `payroll.filter_column`
- `payroll.store_filter_column`
- `payroll.data_range`
- `payroll.headers_range`
- `payroll.benefits.medical.domestic`
- `payroll.benefits.medical.international`
- `payroll.benefits.401k.domestic`
- `payroll.benefits.401k.international`
- `payroll.benefits.other.domestic`
- `payroll.benefits.other.international`
- `payroll.net_new_ARR_achieved`
- `payroll.burn_multiple_achieved`
- `payroll.bonus_cap`
- `payroll.exec_bonus_NNAR_weight`
- `payroll.exec_bonus_burn_multiple_weight`
- `payroll.incentive_bonus_NNAR_weight`
- `payroll.incentive_bonus_burn_multiple_weight`
- `payroll.output.headcount`
- `payroll.output.base_salary_total`
- `payroll.output.base_salary_domestic`
- `payroll.output.base_salary_international`
- `payroll.output.base_salary_cogs`
- `payroll.output.medical`
- `payroll.output.401k`
- `payroll.output.other_benefits`
- `payroll.output.bonus_accrual`
- `payroll.output.bonus_payout`

Example:

```text
Key                                      Value
model.last_actuals_date                  31-Mar-26
model.model_end_date                     30-Apr-28
model.financial_year_end_month           4
payroll.filter_column                    PayrollData!R:R
payroll.store_filter_column              PayrollData!S:S
payroll.data_range                       PayrollData!B5:S1531
payroll.headers_range                    PayrollData!B4:S4
payroll.benefits.medical.domestic        2464
payroll.benefits.medical.international   2162
payroll.benefits.401k.domestic           432
payroll.benefits.401k.international      501
payroll.benefits.other.domestic          157
payroll.benefits.other.international     20
payroll.net_new_ARR_achieved             C_Payroll!BP28:CN28
payroll.burn_multiple_achieved           C_Payroll!BP29:CN29
payroll.bonus_cap                        200%
payroll.exec_bonus_NNAR_weight           75%
payroll.exec_bonus_burn_multiple_weight  25%
payroll.incentive_bonus_NNAR_weight      100%
payroll.incentive_bonus_burn_multiple_weight 0%
payroll.output.headcount                 HCA_Output!E4
payroll.output.base_salary_total         HCA_Output!E17
payroll.output.base_salary_domestic      HCA_Output!E30
payroll.output.base_salary_international HCA_Output!E44
payroll.output.base_salary_cogs          HCA_Output!E57
payroll.output.medical                   HCA_Output!E70
payroll.output.401k                      HCA_Output!E83
payroll.output.other_benefits            HCA_Output!E96
payroll.output.bonus_accrual             HCA_Output!E110
payroll.output.bonus_payout              HCA_Output!E124
```

Headcount/FTE output is written as a table starting at the configured start cell:

```text
Department | Apr 2026 | May 2026 | ...
```

Base salary output uses:

```text
monthly base salary cost = annual salary for financial year / 12 * FTE
```

Benefits output uses:

```text
monthly benefit cost = full monthly benefit assumption when employee FTE is above 0
```

Bonus accrual output uses:

```text
if Bonus $ > 0:
    monthly bonus base = Bonus $ / 12
else:
    monthly bonus base = annual salary for financial year / 12 * Bonus %

bonus accrual = monthly bonus base * plan multiplier * worked-in-month gate * bonus accrual flag
```

Bonus payout output uses:

```text
bonus payout is calculated only in Feb, May, Aug, and Nov
bonus payout = sum of available prior 3 forecast-month bonus accruals
```

## Render Backend

The backend is ready for Render deployment from the repo root using `render.yaml`.

The first MVP uses this temporary environment variable:

```text
CORS_ORIGINS=*
```

Detail storage is optional. To enable it, create a Render Postgres database and add its internal database URL to the backend service:

```text
DATABASE_URL=<Render internal database URL>
```

When `DATABASE_URL` is not set, Payroll Recalc still works and detail storage is skipped.

The backend creates these tables automatically:

```text
calc_runs
calcs_detail_outputs
```

Only the latest run per `User ID` is retained. Each new saved run deletes that user's prior saved run and bulk inserts the new detail rows.

Zero-value detail rows are not saved. Department-level Excel output tables still include zero values where needed.

Detail rows are saved only for employees where the configured `payroll.store_filter_column` equals `1`. This does not affect the department-level calculations written to `HCA_Output`; it only limits what employee-level detail is stored in Postgres.

After a saved recalc, users can load one stored employee/month/output value with:

```excel
=HCA.LOAD_DETAIL("payroll.output.base_salary_total", C1, B10)
```

The arguments are output key, any date in the forecast month, and employee ID. If no stored row exists, the function returns `0`.

The backend response includes separate timing fields for calculation and detail save time. The task pane writes those timings to the activity log after each recalc.

After the add-in is hosted at a stable URL, restrict it to that origin:

```text
CORS_ORIGINS=https://xf1-advisory-services.github.io
```

This MVP has no authentication. Do not treat the preview endpoint as production-secure.

## Microsoft 365 Admin Center Deployment

The production add-in is intended to be hosted on GitHub Pages:

```text
https://xf1-advisory-services.github.io/HCA/taskpane.html
```

Upload this manifest in Microsoft 365 Admin Center:

```text
deploy/m365/hca-production-manifest.xml
```

Recommended first rollout:

1. Deploy to `Just me` first.
2. Test in Windows Desktop Excel.
3. Then assign to a small pilot group.

Microsoft notes that users might need to relaunch Office, and deployed add-ins can take time to appear on the ribbon.
