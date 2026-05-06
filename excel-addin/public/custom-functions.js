(function () {
  const userKeyStorageKey = "xf1.userKey";
  const backendUrlStorageKey = "xf1.backendUrl";
  const defaultBackendUrl = "https://hca-calc-engine.onrender.com";

  async function loadDetail(outputKey, period, unitId) {
    const userKey = normalizeUserKey(await readSharedSetting(userKeyStorageKey));
    if (!userKey) {
      return customFunctionError(
        "Set User ID in the Heavy Calc Assist task pane, then run Payroll Recalc."
      );
    }

    const baseUrl =
      (await readSharedSetting(backendUrlStorageKey)) || defaultBackendUrl;
    const response = await fetch(
      `${baseUrl.replace(/\/$/, "")}/payroll/load-detail`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          userKey,
          outputKey: String(outputKey || "").trim(),
          periodEndDate: normalizePeriodEndDate(period),
          unitId: String(unitId || "").trim(),
        }),
      }
    );

    if (!response.ok) {
      return customFunctionError(`LOAD_DETAIL backend error: ${response.status}`);
    }

    const body = await response.json();
    return Number(body && body.value ? body.value : 0);
  }

  async function readSharedSetting(key) {
    if (window.OfficeRuntime && window.OfficeRuntime.storage) {
      const value = await window.OfficeRuntime.storage.getItem(key);
      if (value) {
        return value;
      }
    }

    return window.localStorage ? window.localStorage.getItem(key) || "" : "";
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

  function customFunctionError(message) {
    if (window.CustomFunctions && window.CustomFunctions.Error) {
      return new window.CustomFunctions.Error(
        window.CustomFunctions.ErrorCode.invalidValue,
        message
      );
    }

    throw new Error(message);
  }

  if (window.CustomFunctions && window.CustomFunctions.associate) {
    window.CustomFunctions.associate("LOAD_DETAIL", loadDetail);
  }
})();
