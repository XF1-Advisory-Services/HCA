import {
  BACKEND_URL_STORAGE_KEY,
  USER_KEY_STORAGE_KEY,
  normalizeUserKey,
  readSharedSetting,
} from "../taskpane/user.js";

const DEFAULT_BACKEND_URL =
  import.meta.env?.VITE_BACKEND_URL || "https://hca-calc-engine.onrender.com";

export async function loadDetail(outputKey, period, unitId) {
  const userKey = normalizeUserKey(await readSharedSetting(USER_KEY_STORAGE_KEY));
  if (!userKey) {
    return customFunctionError(
      "Set User ID in the Heavy Calc Assist task pane, then run Payroll Recalc."
    );
  }

  const baseUrl =
    (await readSharedSetting(BACKEND_URL_STORAGE_KEY)) || DEFAULT_BACKEND_URL;
  const response = await fetch(`${baseUrl.replace(/\/$/, "")}/payroll/load-detail`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      userKey,
      outputKey: String(outputKey ?? "").trim(),
      periodEndDate: normalizePeriodEndDate(period),
      unitId: String(unitId ?? "").trim(),
    }),
  });

  if (!response.ok) {
    return customFunctionError(`LOAD_DETAIL backend error: ${response.status}`);
  }

  return parseLoadDetailValue(await response.json());
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

if (globalThis.CustomFunctions?.associate) {
  globalThis.CustomFunctions.associate("LOAD_DETAIL", loadDetail);
}
