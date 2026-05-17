import logging
import os
import json
from contextlib import asynccontextmanager
from time import perf_counter
from typing import Any
from urllib.parse import parse_qs

from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from pydantic import ValidationError

from app.detail_store import (
    close_detail_store,
    initialize_detail_store,
    load_detail_value,
    load_detail_values,
    save_latest_run,
)
from app.payroll_headcount import calculate_payroll_outputs
from app.schemas import (
    ClientLogRequest,
    PayrollLoadDetailBatchRequest,
    PayrollLoadDetailRequest,
    PayrollLoadPreviewRequest,
)


logger = logging.getLogger("hca.client")


def _cors_origins() -> list[str]:
    raw = os.getenv(
        "CORS_ORIGINS",
        "https://localhost:3000,https://127.0.0.1:3000",
    )
    return [origin.strip() for origin in raw.split(",") if origin.strip()]


@asynccontextmanager
async def lifespan(app: FastAPI):
    try:
        result = initialize_detail_store()
        logger.warning("detail_store_startup status=%s", result)
    except Exception as error:
        logger.exception("detail_store_startup_failed: %s", error)
    yield
    close_detail_store()


app = FastAPI(
    title="XF1 External Calc Engine",
    version="0.1.0",
    lifespan=lifespan,
)

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


@app.get("/payroll/load-detail")
def payroll_load_detail_get(
    userKey: str | None,
    outputKey: str,
    periodEndDate: str,
    unitId: str,
) -> dict[str, Any]:
    return load_detail_value(userKey, outputKey, periodEndDate, unitId)


@app.post("/payroll/load-detail-batch")
def payroll_load_detail_batch(
    payload: PayrollLoadDetailBatchRequest,
) -> dict[str, Any]:
    started_at = perf_counter()
    result = load_detail_values(payload.userKey, payload.items)
    timings = {
        "totalBackendMs": elapsed_ms(started_at),
        "itemCount": len(payload.items),
    }
    result["timings"] = timings
    logger.warning(
        "load_detail_batch status=%s items=%s found=%s total_ms=%s",
        result.get("status"),
        len(payload.items),
        result.get("foundCount"),
        timings["totalBackendMs"],
    )
    return result


@app.post("/payroll/load-detail-batch-text")
async def payroll_load_detail_batch_text(request: Request) -> dict[str, Any]:
    body_text = (await request.body()).decode("utf-8")
    try:
        payload = parse_load_detail_batch_transport(
            request.headers.get("content-type", ""),
            body_text,
        )
    except ValueError as error:
        raise HTTPException(status_code=400, detail=str(error)) from error

    return payroll_load_detail_batch(payload)


def parse_load_detail_batch_transport(
    content_type: str,
    body_text: str,
) -> PayrollLoadDetailBatchRequest:
    normalized_content_type = content_type.split(";", 1)[0].strip().lower()

    if normalized_content_type == "application/x-www-form-urlencoded":
        form_values = parse_qs(body_text, keep_blank_values=True)
        payload_values = form_values.get("payload")
        if not payload_values:
            raise ValueError("Missing form field: payload")
        raw_payload = payload_values[0]
    else:
        raw_payload = body_text

    try:
        decoded = json.loads(raw_payload)
    except json.JSONDecodeError as error:
        raise ValueError("Batch payload must be valid JSON") from error

    try:
        return PayrollLoadDetailBatchRequest.model_validate(decoded)
    except ValidationError as error:
        raise ValueError(str(error)) from error


@app.post("/debug/client-log")
def debug_client_log(payload: ClientLogRequest) -> dict[str, str]:
    logger.warning(
        "client_log source=%s stage=%s level=%s message=%s context=%s",
        payload.source,
        payload.stage,
        payload.level,
        payload.message,
        payload.context,
    )
    return {"status": "logged"}


def elapsed_ms(started_at: float) -> float:
    return round((perf_counter() - started_at) * 1000, 2)
