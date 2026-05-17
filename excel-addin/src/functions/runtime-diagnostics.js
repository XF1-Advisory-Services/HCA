const DEFAULT_BACKEND_URL = "https://hca-calc-engine.onrender.com";
const runtimeId = Math.random().toString(36).slice(2, 10);
let batchSequence = 0;
const diagnosticBatchGroups = new Map();

export async function runtimeCheck(options = {}) {
  const setTimeoutFn =
    Object.prototype.hasOwnProperty.call(options, "setTimeoutFn")
      ? options.setTimeoutFn
      : globalThis.setTimeout;
  const fetchFn =
    Object.prototype.hasOwnProperty.call(options, "fetchFn")
      ? options.fetchFn
      : globalThis.fetch;
  const storage =
    Object.prototype.hasOwnProperty.call(options, "storage")
      ? options.storage
      : globalThis.OfficeRuntime?.storage;

  const promiseStatus = await Promise.resolve("ok");
  const timerStatus =
    typeof setTimeoutFn === "function"
      ? await new Promise((resolve) => {
          setTimeoutFn(() => resolve("ok"), 0);
        })
      : "missing";

  return [
    `runtime=${runtimeId}`,
    `promise=${promiseStatus}`,
    `setTimeout=${typeof setTimeoutFn === "function" ? "function" : "missing"}`,
    `timer=${timerStatus}`,
    `fetch=${typeof fetchFn === "function" ? "function" : "missing"}`,
    `officeStorage=${storage ? "available" : "missing"}`,
  ].join(";");
}

const POST_CHECK_MODES = new Set(["json", "text", "empty", "form"]);

export async function postCheck(modeOrBaseUrl = "json", baseUrlOrOptions = "", options = {}) {
  const { mode, baseUrl, resolvedOptions } = resolvePostCheckArgs(
    modeOrBaseUrl,
    baseUrlOrOptions,
    options
  );
  const fetchFn = resolvedOptions.fetchFn ?? fetch;
  const cleanBaseUrl = normalizeBaseUrl(baseUrl);

  try {
    const response = await fetchFn(
      `${cleanBaseUrl}/debug/client-log`,
      buildPostCheckRequest(mode)
    );

    return `mode=${mode};post=${Number(response.status || 0)}`;
  } catch (error) {
    return `mode=${mode};post=error;message=${truncate(String(error?.message || error))}`;
  }
}

export function batchQueueCheck(token = "default", delayMs = 100, options = {}) {
  const cleanToken = String(token || "default").trim() || "default";
  const cleanDelayMs = Math.max(0, Number(delayMs || 0));

  let group = diagnosticBatchGroups.get(cleanToken);
  if (!group) {
    group = {
      id: ++batchSequence,
      token: cleanToken,
      entries: [],
      scheduler: "none",
      timerId: null,
    };
    diagnosticBatchGroups.set(cleanToken, group);
  }

  const promise = new Promise((resolve) => {
    group.entries.push({ resolve });
  });

  if (!group.timerId) {
    group.timerId = scheduleDiagnosticFlush(
      () => flushDiagnosticBatch(cleanToken),
      cleanDelayMs,
      options
    );
    group.scheduler = group.timerId === true ? "microtask" : "setTimeout";
  }

  return promise;
}

function flushDiagnosticBatch(token) {
  const group = diagnosticBatchGroups.get(token);
  if (!group) {
    return;
  }

  diagnosticBatchGroups.delete(token);
  const size = group.entries.length;
  group.entries.forEach((entry, index) => {
    entry.resolve(
      [
        `runtime=${runtimeId}`,
        `token=${group.token}`,
        `batch=${group.id}`,
        `index=${index + 1}`,
        `size=${size}`,
        `scheduler=${group.scheduler}`,
      ].join(";")
    );
  });
}

function scheduleDiagnosticFlush(callback, delayMs, options = {}) {
  if (typeof options.setTimeoutFn === "function") {
    return options.setTimeoutFn(callback, delayMs);
  }

  if (typeof globalThis.setTimeout === "function") {
    return globalThis.setTimeout(callback, delayMs);
  }

  Promise.resolve().then(callback);
  return true;
}

function normalizeBaseUrl(baseUrl) {
  return String(baseUrl || DEFAULT_BACKEND_URL).trim().replace(/\/$/, "");
}

function resolvePostCheckArgs(modeOrBaseUrl, baseUrlOrOptions, options) {
  const first = String(modeOrBaseUrl || "json").trim();
  const firstLower = first.toLowerCase();

  if (POST_CHECK_MODES.has(firstLower)) {
    return {
      mode: firstLower,
      baseUrl: typeof baseUrlOrOptions === "string" ? baseUrlOrOptions : "",
      resolvedOptions:
        typeof baseUrlOrOptions === "object" && baseUrlOrOptions !== null
          ? baseUrlOrOptions
          : options,
    };
  }

  return {
    mode: "json",
    baseUrl: first,
    resolvedOptions:
      typeof baseUrlOrOptions === "object" && baseUrlOrOptions !== null
        ? baseUrlOrOptions
        : options,
  };
}

function buildPostCheckRequest(mode) {
  if (mode === "empty") {
    return { method: "POST" };
  }

  if (mode === "text") {
    return {
      method: "POST",
      headers: {
        "Content-Type": "text/plain",
      },
      body: `source=HCA.POST_CHECK;runtime=${runtimeId}`,
    };
  }

  if (mode === "form") {
    return {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: new URLSearchParams({
        source: "HCA.POST_CHECK",
        runtime: runtimeId,
      }).toString(),
    };
  }

  return {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      source: "HCA.POST_CHECK",
      stage: "post-check",
      level: "info",
      message: "Custom function POST probe.",
      context: { runtimeId, mode },
    }),
  };
}

function truncate(value) {
  return value.length > 120 ? `${value.slice(0, 120)}...` : value;
}
