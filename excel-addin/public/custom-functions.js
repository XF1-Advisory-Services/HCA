(function () {
  const root = globalThis;
  const userKeyStorageKey = "xf1.userKey";
  const backendUrlStorageKey = "xf1.backendUrl";
  const defaultBackendUrl = "https://hca-calc-engine.onrender.com";
  const loadDetailBatchDelayMs = 50;
  const loadDetailBatchMaxSize = 500;
  const settingsCacheTtlMs = 1000;
  const lookupKeyDelimiter = "\u001f";

  let settingsCache = {
    expiresAt: 0,
    promise: null,
  };
  const batchGroups = new Map();
  const pendingByLookupKey = new Map();

  async function loadDetail(outputKey, period, unitId, userKeyOverride) {
    let stage = "start";
    let baseUrl = defaultBackendUrl;
    const context = buildClientLogContext(outputKey, period, unitId, userKeyOverride);

    try {
      stage = "read-settings";
      const settings = await readLoadDetailSettings();
      const userKey = normalizeUserKey(userKeyOverride || settings.userKey);
      context.userKey = userKey;
      if (!userKey) {
        const message =
          "Set User ID in the Heavy Calc Assist task pane, then run Payroll Recalc.";
        await reportClientError(baseUrl, "missing-user-key", message, context);
        return customFunctionError(message);
      }

      stage = "read-backend-url";
      baseUrl = settings.backendUrl || defaultBackendUrl;
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

      stage = "backend-batch";
      return await queueLoadDetailLookup(baseUrl, userKey, {
        outputKey: requestBody.outputKey,
        periodEndDate: requestBody.periodEndDate,
        unitId: requestBody.unitId,
      });
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
      const response = await fetch(
        buildLoadDetailUrl(defaultBackendUrl, {
          userKey: "vavrinec@xf1advisory.com",
          outputKey: "payroll.output.401k",
          periodEndDate: "2026-04-30",
          unitId: "EX18",
        })
      );
      if (!response.ok) {
        return -Number(response.status || 1);
      }
      const body = await response.json();
      return Number(body && body.value ? body.value : 0);
    } catch {
      return -1;
    }
  }

  function buildLoadDetailUrl(baseUrl, requestBody) {
    const params = new URLSearchParams({
      userKey: requestBody.userKey,
      outputKey: requestBody.outputKey,
      periodEndDate: requestBody.periodEndDate,
      unitId: requestBody.unitId,
    });
    return `${baseUrl.replace(/\/$/, "")}/payroll/load-detail?${params.toString()}`;
  }

  function buildLoadDetailLookupKey(userKey, item) {
    return [
      normalizeUserKey(userKey),
      String(item.outputKey || "").trim(),
      String(item.periodEndDate || "").trim(),
      String(item.unitId || "").trim(),
    ].join(lookupKeyDelimiter);
  }

  function queueLoadDetailLookup(baseUrl, userKey, item, options) {
    const cleanBaseUrl = normalizeBaseUrl(baseUrl || defaultBackendUrl);
    const cleanUserKey = normalizeUserKey(userKey);
    const groupKey = buildLoadDetailGroupKey(cleanBaseUrl, cleanUserKey);
    const itemLookupKey = buildLoadDetailLookupKey(cleanUserKey, item);
    const pendingKey = `${groupKey}${lookupKeyDelimiter}${itemLookupKey}`;

    if (pendingByLookupKey.has(pendingKey)) {
      return pendingByLookupKey.get(pendingKey).promise;
    }

    let resolve;
    let reject;
    const promise = new Promise((res, rej) => {
      resolve = res;
      reject = rej;
    });

    pendingByLookupKey.set(pendingKey, { promise });

    let group = batchGroups.get(groupKey);
    if (!group) {
      group = {
        baseUrl: cleanBaseUrl,
        userKey: cleanUserKey,
        items: [],
        timerId: null,
        options: options || {},
      };
      batchGroups.set(groupKey, group);
    }

    group.items.push({
      lookupKey: pendingKey,
      item: {
        outputKey: String(item.outputKey || "").trim(),
        periodEndDate: String(item.periodEndDate || "").trim(),
        unitId: String(item.unitId || "").trim(),
      },
      resolve,
      reject,
    });

    const delayMs =
      options && options.delayMs !== undefined
        ? options.delayMs
        : loadDetailBatchDelayMs;
    if (!group.timerId) {
      group.timerId = scheduleLoadDetailFlush(() => {
        flushLoadDetailGroup(groupKey);
      }, delayMs, options || {});
    }

    return promise;
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

  async function readLoadDetailSettings() {
    const now = Date.now();
    if (settingsCache.promise && settingsCache.expiresAt > now) {
      return settingsCache.promise;
    }

    settingsCache.promise = Promise.all([
      readSharedSetting(userKeyStorageKey),
      readSharedSetting(backendUrlStorageKey),
    ]).then(([userKey, backendUrl]) => ({
      userKey: normalizeUserKey(userKey),
      backendUrl: backendUrl || defaultBackendUrl,
    }));

    settingsCache.expiresAt = now + settingsCacheTtlMs;
    return settingsCache.promise;
  }

  async function flushLoadDetailGroup(groupKey) {
    const group = batchGroups.get(groupKey);
    if (!group) {
      return;
    }

    batchGroups.delete(groupKey);
    group.timerId = null;

    const maxBatchSize = group.options.maxBatchSize || loadDetailBatchMaxSize;
    const chunks = chunkArray(group.items, maxBatchSize);

    await Promise.all(
      chunks.map((chunk) =>
        sendLoadDetailBatch(group.baseUrl, group.userKey, chunk, group.options)
      )
    );
  }

  function chunkArray(values, size) {
    const chunks = [];
    for (let index = 0; index < values.length; index += size) {
      chunks.push(values.slice(index, index + size));
    }
    return chunks;
  }

  async function sendLoadDetailBatch(baseUrl, userKey, queuedItems, options) {
    const fetchFn = options && options.fetchFn ? options.fetchFn : fetch;

    try {
      const response = await fetchFn(`${baseUrl}/payroll/load-detail-batch`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          userKey,
          items: queuedItems.map(({ item }) => ({
            outputKey: item.outputKey,
            periodEndDate: item.periodEndDate,
            unitId: item.unitId,
          })),
        }),
      });

      if (!response.ok) {
        const responseText = await readResponseText(response);
        throw new Error(
          `LOAD_DETAIL batch backend error: ${response.status} ${responseText}`
        );
      }

      const body = await response.json();
      const values = Array.isArray(body && body.values) ? body.values : [];

      if (values.length !== queuedItems.length) {
        throw new Error(
          `LOAD_DETAIL batch returned ${values.length} values for ${queuedItems.length} lookups.`
        );
      }

      queuedItems.forEach((queuedItem, index) => {
        pendingByLookupKey.delete(queuedItem.lookupKey);
        queuedItem.resolve(Number(values[index] || 0));
      });
    } catch (error) {
      queuedItems.forEach((queuedItem) => {
        pendingByLookupKey.delete(queuedItem.lookupKey);
        queuedItem.reject(error);
      });
    }
  }

  function buildLoadDetailGroupKey(baseUrl, userKey) {
    return `${baseUrl}${lookupKeyDelimiter}${userKey}`;
  }

  function normalizeBaseUrl(baseUrl) {
    return String(baseUrl || defaultBackendUrl).trim().replace(/\/$/, "");
  }

  function scheduleLoadDetailFlush(callback, delayMs, options) {
    if (options && typeof options.setTimeoutFn === "function") {
      return options.setTimeoutFn(callback, delayMs);
    }

    if (typeof root.setTimeout === "function") {
      return root.setTimeout(callback, delayMs);
    }

    Promise.resolve().then(callback);
    return true;
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
