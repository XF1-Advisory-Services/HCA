import {
  BACKEND_URL_STORAGE_KEY,
  USER_KEY_STORAGE_KEY,
  normalizeUserKey,
  readSharedSetting,
} from "../taskpane/user.js";

const DEFAULT_BACKEND_URL =
  import.meta.env?.VITE_BACKEND_URL || "https://hca-calc-engine.onrender.com";

export async function loadDetail(outputKey, period, unitId, userKeyOverride = "") {
  let stage = "start";
  let baseUrl = DEFAULT_BACKEND_URL;
  const context = buildClientLogContext(outputKey, period, unitId, userKeyOverride);

  try {
    stage = "read-user-key";
    const userKey = normalizeUserKey(
      userKeyOverride || (await readSharedSetting(USER_KEY_STORAGE_KEY))
    );
    context.userKey = userKey;
    if (!userKey) {
      const message =
        "Set User ID in the Heavy Calc Assist task pane, then run Payroll Recalc.";
      await reportClientError(baseUrl, "missing-user-key", message, context);
      return customFunctionError(message);
    }

    stage = "read-backend-url";
    baseUrl = (await readSharedSetting(BACKEND_URL_STORAGE_KEY)) || DEFAULT_BACKEND_URL;
    context.backendUrl = baseUrl;

    stage = "normalize-period";
    const requestBody = {
      userKey,
      outputKey: String(outputKey ?? "").trim(),
      periodEndDate: normalizePeriodEndDate(period),
      unitId: String(unitId ?? "").trim(),
    };
    context.outputKey = requestBody.outputKey;
    context.periodEndDate = requestBody.periodEndDate;
    context.unitId = requestBody.unitId;

    stage = "backend-fetch";
    const response = await fetch(buildLoadDetailUrl(baseUrl, requestBody));

    context.responseStatus = response.status;
    if (!response.ok) {
      const responseText = await readResponseText(response);
      await reportClientError(
        baseUrl,
        "backend-response",
        `LOAD_DETAIL backend error: ${response.status}`,
        { ...context, responseText }
      );
      return customFunctionError(`LOAD_DETAIL backend error: ${response.status}`);
    }

    stage = "backend-json";
    const responseBody = await response.json();

    stage = "parse-value";
    return parseLoadDetailValue(responseBody);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    await reportClientError(baseUrl, stage, message, context);
    return customFunctionError(
      `LOAD_DETAIL failed at ${stage}. Check Render logs for details.`
    );
  }
}

export function diag() {
  return 123;
}

export async function diagBackend() {
  const baseUrl =
    (await readSharedSetting(BACKEND_URL_STORAGE_KEY)) || DEFAULT_BACKEND_URL;
  return getBackendHealthStatus(baseUrl);
}

export async function diagLoadDetail(fetchFn = fetch) {
  try {
    const response = await fetchFn(
      buildLoadDetailUrl(DEFAULT_BACKEND_URL, {
        userKey: "vavrinec@xf1advisory.com",
        outputKey: "payroll.output.401k",
        periodEndDate: "2026-04-30",
        unitId: "EX18",
      })
    );
    if (!response.ok) {
      return -Number(response.status || 1);
    }
    return parseLoadDetailValue(await response.json());
  } catch {
    return -1;
  }
}

export function buildLoadDetailUrl(baseUrl, requestBody) {
  const params = new URLSearchParams({
    userKey: requestBody.userKey,
    outputKey: requestBody.outputKey,
    periodEndDate: requestBody.periodEndDate,
    unitId: requestBody.unitId,
  });
  return `${baseUrl.replace(/\/$/, "")}/payroll/load-detail?${params.toString()}`;
}

export async function getBackendHealthStatus(baseUrl, fetchFn = fetch) {
  try {
    const response = await fetchFn(`${baseUrl.replace(/\/$/, "")}/health`);
    return Number(response.status || 0);
  } catch {
    return -1;
  }
}

export function normalizePeriodEndDate(value) {
  const date = parseInputDate(value);
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 0))
    .toISOString()
    .slice(0, 10);
}

export function parseLoadDetailValue(responseBody) {
  return Number(responseBody?.value || 0);
}

export function buildClientLogContext(outputKey, period, unitId, userKeyOverride = "") {
  return {
    outputKey: String(outputKey ?? "").trim(),
    periodType: typeof period,
    periodRaw: safeDebugValue(period),
    unitId: String(unitId ?? "").trim(),
    userKeyOverrideProvided: Boolean(String(userKeyOverride ?? "").trim()),
  };
}

function parseInputDate(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return new Date(Date.UTC(1899, 11, 30 + Math.floor(value)));
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(
      Date.UTC(value.getFullYear(), value.getMonth(), value.getDate())
    );
  }

  const parsed = new Date(value);
  if (!Number.isNaN(parsed.getTime())) {
    return new Date(
      Date.UTC(parsed.getFullYear(), parsed.getMonth(), parsed.getDate())
    );
  }

  throw new Error("Period must be a valid Excel date.");
}

function customFunctionError(message) {
  if (globalThis.CustomFunctions?.Error) {
    return new globalThis.CustomFunctions.Error(
      globalThis.CustomFunctions.ErrorCode.invalidValue,
      message
    );
  }

  throw new Error(message);
}

async function reportClientError(baseUrl, stage, message, context) {
  try {
    await fetch(`${baseUrl.replace(/\/$/, "")}/debug/client-log`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        source: "HCA.LOAD_DETAIL",
        stage,
        level: "error",
        message,
        context: scrubClientLogContext(context),
      }),
    });
  } catch {
    // Diagnostics must never create a second worksheet error.
  }
}

async function readResponseText(response) {
  try {
    return truncateDebugText(await response.text());
  } catch {
    return "";
  }
}

function scrubClientLogContext(context) {
  const clean = {};
  for (const [key, value] of Object.entries(context || {})) {
    clean[key] = safeDebugValue(value);
  }
  return clean;
}

function safeDebugValue(value) {
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? "Invalid Date" : value.toISOString();
  }

  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "object") {
    return truncateDebugText(JSON.stringify(value));
  }

  return truncateDebugText(String(value));
}

function truncateDebugText(value) {
  const text = String(value ?? "");
  return text.length > 500 ? `${text.slice(0, 500)}...` : text;
}

if (globalThis.CustomFunctions?.associate) {
  globalThis.CustomFunctions.associate("LOAD_DETAIL", loadDetail);
  globalThis.CustomFunctions.associate("DIAG", diag);
  globalThis.CustomFunctions.associate("DIAG_BACKEND", diagBackend);
  globalThis.CustomFunctions.associate("DIAG_LOAD_DETAIL", diagLoadDetail);
}
