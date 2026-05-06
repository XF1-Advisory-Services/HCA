import os
from time import perf_counter
from typing import Any

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.detail_store import load_detail_value, save_latest_run
from app.payroll_headcount import calculate_payroll_outputs
from app.schemas import PayrollLoadDetailRequest, PayrollLoadPreviewRequest


def _cors_origins() -> list[str]:
    raw = os.getenv(
        "CORS_ORIGINS",
        "https://localhost:3000,https://127.0.0.1:3000",
    )
    return [origin.strip() for origin in raw.split(",") if origin.strip()]


app = FastAPI(title="XF1 External Calc Engine", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=_cors_origins(),
    allow_credentials=False,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/payroll/load-preview")
def payroll_load_preview(payload: PayrollLoadPreviewRequest) -> dict[str, Any]:
    backend_started_at = perf_counter()
    sample_keys = list(payload.rows[0].keys()) if payload.rows else []
    calculation_started_at = perf_counter()
    outputs = calculate_payroll_outputs(
        payload.headers, payload.rows, payload.model, payload.assumptions
    )
    calculation_ms = elapsed_ms(calculation_started_at)
    detail_rows = outputs.pop("detailRows", [])
    detail_save_started_at = perf_counter()
    try:
        detail_save = save_latest_run(payload.userKey, payload, detail_rows)
    except Exception as error:
        detail_save = {
            "status": "error",
            "reason": str(error),
            "rowsPrepared": len(detail_rows),
            "rowsSaved": 0,
        }
    detail_save_ms = elapsed_ms(detail_save_started_at)

    return {
        "status": "received",
        "detailSave": detail_save,
        "timings": {
            "calculationMs": calculation_ms,
            "detailSaveMs": detail_save_ms,
            "totalBackendMs": elapsed_ms(backend_started_at),
        },
        "section": payload.section,
        "model": {
            "lastActualsDate": payload.model.lastActualsDate,
            "modelEndDate": payload.model.modelEndDate,
            "calculationStartDate": payload.model.calculationStartDate,
            "calculationEndDate": payload.model.calculationEndDate,
            "calculationMonths": payload.model.calculationMonths,
            "financialYearEndMonth": payload.model.financialYearEndMonth,
            "firstPeriod": payload.model.periods[0].model_dump() if payload.model.periods else None,
            "lastPeriod": payload.model.periods[-1].model_dump() if payload.model.periods else None,
        },
        "source": payload.source.model_dump(),
        "totalRows": payload.metrics.totalRows,
        "includedRows": payload.metrics.includedRows,
        "headersCount": len(payload.headers),
        "receivedRows": len(payload.rows),
        "loadTimeMs": payload.metrics.loadTimeMs,
        "sampleKeys": sample_keys[:10],
        "outputs": outputs,
    }


@app.post("/payroll/load-detail")
def payroll_load_detail(payload: PayrollLoadDetailRequest) -> dict[str, Any]:
    return load_detail_value(
        payload.userKey,
        payload.outputKey,
        payload.periodEndDate,
        payload.unitId,
    )


def elapsed_ms(started_at: float) -> float:
    return round((perf_counter() - started_at) * 1000, 2)
