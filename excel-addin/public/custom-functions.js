(function () {
  const root = globalThis;
  const userKeyStorageKey = "xf1.userKey";
  const backendUrlStorageKey = "xf1.backendUrl";
  const defaultBackendUrl = "https://hca-calc-engine.onrender.com";

  async function loadDetail(outputKey, period, unitId, userKeyOverride) {
    let stage = "start";
    let baseUrl = defaultBackendUrl;
    const context = buildClientLogContext(outputKey, period, unitId, userKeyOverride);

    try {
      stage = "read-user-key";
      const userKey = normalizeUserKey(
        userKeyOverride || (await readSharedSetting(userKeyStorageKey))
      );
      context.userKey = userKey;
      if (!userKey) {
        const message =
          "Set User ID in the Heavy Calc Assist task pane, then run Payroll Recalc.";
        await reportClientError(baseUrl, "missing-user-key", message, context);
        return customFunctionError(message);
      }

      stage = "read-backend-url";
      baseUrl = (await readSharedSetting(backendUrlStorageKey)) || defaultBackendUrl;
      context.backendUrl = baseUrl;

      stage = "normalize-period";
      const requestBody = {
        userKey,
        outputKey: String(outputKey || "").trim(),
        periodEndDate: normalizePeriodEndDate(period),
        unitId: String(unitId || "").trim(),
      };
      context.outputKey = requestBody.outputKey;
      context.periodEndDate = requestBody.periodEndDate;
      context.unitId = requestBody.unitId;

      stage = "backend-fetch";
      const response = await fetch(`${baseUrl.replace(/\/$/, "")}/payroll/load-detail`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(requestBody),
      });

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
      const body = await response.json();

      stage = "parse-value";
      return Number(body && body.value ? body.value : 0);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      await reportClientError(baseUrl, stage, message, context);
      return customFunctionError(
        `LOAD_DETAIL failed at ${stage}. Check Render logs for details.`
      );
    }
  }

  function diag() {
    return 123;
  }

  async function diagBackend() {
    const baseUrl =
      (await readSharedSetting(backendUrlStorageKey)) || defaultBackendUrl;
    return getBackendHealthStatus(baseUrl);
  }

  async function diagLoadDetail() {
    try {
      const response = await fetch(`${defaultBackendUrl}/payroll/load-detail`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          userKey: "vavrinec@xf1advisory.com",
          outputKey: "payroll.output.401k",
          periodEndDate: "2026-04-30",
          unitId: "EX18",
        }),
      });
      if (!response.ok) {
        return -Number(response.status || 1);
      }
      const body = await response.json();
      return Number(body && body.value ? body.value : 0);
    } catch {
      return -1;
    }
  }

  async function getBackendHealthStatus(baseUrl, fetchFn) {
    try {
      const response = await (fetchFn || fetch)(`${baseUrl.replace(/\/$/, "")}/health`);
      return Number(response.status || 0);
    } catch {
      return -1;
    }
  }

  async function readSharedSetting(key) {
    if (root.OfficeRuntime && root.OfficeRuntime.storage) {
      const value = await root.OfficeRuntime.storage.getItem(key);
      if (value) {
        return value;
      }
    }

    return root.localStorage ? root.localStorage.getItem(key) || "" : "";
  }

  function normalizePeriodEndDate(value) {
    const date = parseInputDate(value);
    return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 0))
      .toISOString()
      .slice(0, 10);
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

  function normalizeUserKey(value) {
    return String(value || "").trim().toLowerCase();
  }

  function buildClientLogContext(outputKey, period, unitId, userKeyOverride) {
    return {
      outputKey: String(outputKey || "").trim(),
      periodType: typeof period,
      periodRaw: safeDebugValue(period),
      unitId: String(unitId || "").trim(),
      userKeyOverrideProvided: Boolean(String(userKeyOverride || "").trim()),
    };
  }

  function customFunctionError(message) {
    if (root.CustomFunctions && root.CustomFunctions.Error) {
      return new root.CustomFunctions.Error(
        root.CustomFunctions.ErrorCode.invalidValue,
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
    const text = String(value || "");
    return text.length > 500 ? `${text.slice(0, 500)}...` : text;
  }

  if (root.CustomFunctions && root.CustomFunctions.associate) {
    root.CustomFunctions.associate("LOAD_DETAIL", loadDetail);
    root.CustomFunctions.associate("DIAG", diag);
    root.CustomFunctions.associate("DIAG_BACKEND", diagBackend);
    root.CustomFunctions.associate("DIAG_LOAD_DETAIL", diagLoadDetail);
  }
})();
