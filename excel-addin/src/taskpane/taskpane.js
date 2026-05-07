import "./taskpane.css";
import {
  getConfigNamedRange,
  getFilterOffset,
  isSelectedFlag,
  parseConfig,
} from "./config.js";
import {
  BACKEND_URL_STORAGE_KEY,
  USER_KEY_STORAGE_KEY,
  normalizeUserKey,
  readSharedSetting,
  writeSharedSetting,
} from "./user.js";

const DEFAULT_BACKEND_URL =
  import.meta.env.VITE_BACKEND_URL || "https://hca-calc-engine.onrender.com";

const elements = {
  button: document.querySelector("#payroll-recalc"),
  clearLog: document.querySelector("#clear-log"),
  status: document.querySelector("#status"),
  configState: document.querySelector("#config-state"),
  dataSheet: document.querySelector("#data-sheet"),
  dataRange: document.querySelector("#data-range"),
  filterRule: document.querySelector("#filter-rule"),
  forecastPeriod: document.querySelector("#forecast-period"),
  calculationMonths: document.querySelector("#calculation-months"),
  includedRows: document.querySelector("#included-rows"),
  loadSpeed: document.querySelector("#load-speed"),
  backendResult: document.querySelector("#backend-result"),
  activityLog: document.querySelector("#activity-log"),
  backendUrl: document.querySelector("#backend-url"),
  userKey: document.querySelector("#user-key"),
  backendState: document.querySelector("#backend-state"),
  cloudDot: document.querySelector("#cloud-dot"),
};

Office.onReady(async (info) => {
  elements.backendUrl.value =
    (await readSharedSetting(BACKEND_URL_STORAGE_KEY)) || DEFAULT_BACKEND_URL;
  elements.userKey.value = await readSharedSetting(USER_KEY_STORAGE_KEY);
  elements.backendState.textContent = safeHost(elements.backendUrl.value);
  addLog("Task pane ready.");

  if (info.host !== Office.HostType.Excel) {
    setStatus("Open this add-in inside Excel Desktop.", "error");
    elements.button.disabled = true;
    return;
  }

  elements.button.addEventListener("click", handlePayrollRecalc);
  elements.clearLog.addEventListener("click", () => {
    elements.activityLog.innerHTML = "";
    addLog("Activity log cleared.");
  });
  elements.backendUrl.addEventListener("change", () => {
    const value = elements.backendUrl.value.trim();
    writeSharedSetting(BACKEND_URL_STORAGE_KEY, value);
    elements.backendState.textContent = value ? safeHost(value) : "Backend";
    checkBackendHealth();
  });
  elements.userKey.addEventListener("change", () => {
    elements.userKey.value = normalizeUserKey(elements.userKey.value);
    writeSharedSetting(USER_KEY_STORAGE_KEY, elements.userKey.value);
  });

  checkBackendHealth();
});

async function handlePayrollRecalc() {
  const startedAt = performance.now();
  setBusy(true);
  setStatus("Reading Payroll data from workbook...");
  addLog("Payroll Recalc started.");

  try {
    const payload = await buildPayrollPayload(startedAt);
    updateConfigUi(payload.source, payload.model);
    updateMetrics(payload.metrics);
    await clearOutputRange(payload.output);
    setStatus("Workbook load complete. Sending preview to backend...");

    const backendSummary = await sendLoadPreview(payload);
    await writePayrollOutputs(payload.output, backendSummary.outputs);
    reportBackendTimings(backendSummary.timings);
    reportDetailSave(backendSummary.detailSave);
    setBackendUi("Success", "connected");
    elements.backendResult.textContent = backendSummary.status || "Success";
    elements.backendResult.className = "is-success";
    setStatus("Payroll outputs complete.", "success");
    addLog(
      `Payroll outputs completed. ${payload.metrics.includedRows.toLocaleString()} rows included.`
    );
  } catch (error) {
    setBackendUi("Error", "error");
    elements.backendResult.textContent = "Error";
    elements.backendResult.className = "is-error";
    setStatus(error.message || String(error), "error");
    addLog(error.message || String(error));
  } finally {
    setBusy(false);
  }
}

async function buildPayrollPayload(startedAt) {
  return Excel.run(async (context) => {
    const configRange = getConfigNamedRange(context);
    configRange.load("values");

    await context.sync();

    const config = parseConfig(configRange.values);
    const dataSheet = context.workbook.worksheets.getItem(config.payroll.dataLoadSheet);
    const headerRange = dataSheet.getRange(config.payroll.headers);
    const dataRange = dataSheet.getRange(config.payroll.cellRange);
    const netNewArrRange = getReferencedRange(
      context,
      config.seriesRanges.netNewArrAchieved
    );
    const burnMultipleRange = getReferencedRange(
      context,
      config.seriesRanges.burnMultipleAchieved
    );

    headerRange.load("values");
    dataRange.load("values");
    netNewArrRange.load("values");
    burnMultipleRange.load("values");

    await context.sync();

    const headers = normalizeHeaders(headerRange.values[0] || []);
    const rows = dataRange.values || [];
    const filterOffset = getFilterOffset(
      config.payroll.cellRange,
      config.payroll.filterColumn
    );
    const storeFilterOffset = getFilterOffset(
      config.payroll.cellRange,
      config.payroll.storeFilterColumn
    );
    const included = rows
      .filter((row) => isSelectedFlag(row[filterOffset]))
      .map((row) => ({
        ...rowToObject(headers, row),
        __hcaStoreDetail: isSelectedFlag(row[storeFilterOffset]),
      }));
    const storeRows = included.filter((row) => row.__hcaStoreDetail).length;

    const loadTimeMs = Math.round((performance.now() - startedAt) * 10) / 10;

    return {
      section: "Payroll",
      userKey: normalizeUserKey(elements.userKey.value),
      model: config.model,
      source: {
        sheet: config.payroll.dataLoadSheet,
        headerRange: config.payroll.headers,
        dataRange: config.payroll.cellRange,
        filterColumn: config.payroll.filterColumn,
        storeFilterColumn: config.payroll.storeFilterColumn,
      },
      output: config.output,
      assumptions: {
        ...config.assumptions,
        bonus: {
          ...config.assumptions.bonus,
          netNewArrAchieved: flattenRangeValues(netNewArrRange.values),
          burnMultipleAchieved: flattenRangeValues(burnMultipleRange.values),
        },
      },
      metrics: {
        totalRows: rows.length,
        includedRows: included.length,
        storeRows,
        loadTimeMs,
      },
      headers,
      rows: included,
    };
  });
}

function getReferencedRange(context, reference) {
  return context.workbook.worksheets
    .getItem(reference.sheet)
    .getRange(reference.address);
}

function flattenRangeValues(values) {
  return (values || []).flat().filter((value) => value !== null && value !== "");
}

function normalizeHeaders(headers) {
  return headers.map((header, index) => {
    const normalized = String(header ?? "").trim();
    return normalized || `Column ${index + 1}`;
  });
}

function rowToObject(headers, row) {
  return headers.reduce((record, header, index) => {
    record[header] = row[index] ?? null;
    return record;
  }, {});
}

async function sendLoadPreview(payload) {
  const baseUrl = elements.backendUrl.value.trim().replace(/\/$/, "");
  await writeSharedSetting(USER_KEY_STORAGE_KEY, normalizeUserKey(payload.userKey));
  await writeSharedSetting(BACKEND_URL_STORAGE_KEY, baseUrl);
  elements.backendState.textContent = safeHost(baseUrl);

  const response = await fetch(`${baseUrl}/payroll/load-preview`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Backend returned ${response.status}: ${body}`);
  }

  return response.json();
}

function reportDetailSave(detailSave) {
  if (!detailSave) {
    return;
  }

  if (detailSave.status === "saved") {
    addLog(
      `Saved ${Number(detailSave.rowsSaved || 0).toLocaleString()} detail rows for latest run.`
    );
    return;
  }

  if (detailSave.status === "skipped") {
    addLog(`Detail storage skipped: ${detailSave.reason}.`);
    return;
  }

  addLog(`Detail storage error: ${detailSave.reason || "unknown error"}.`);
}

function reportBackendTimings(timings) {
  if (!timings) {
    return;
  }

  addLog(
    `Backend timing: calc ${formatMs(timings.calculationMs)}, detail save ${formatMs(
      timings.detailSaveMs
    )}.`
  );
}

function formatMs(value) {
  const milliseconds = Number(value || 0);
  if (milliseconds >= 1000) {
    return `${(milliseconds / 1000).toFixed(2)}s`;
  }
  return `${Math.round(milliseconds)}ms`;
}

async function clearOutputRange(outputConfig) {
  const cleared = await Excel.run(async (context) => {
    const outputSheet = context.workbook.worksheets.getItem(outputConfig.sheet);
    const usedRange = outputSheet.getUsedRangeOrNullObject();
    usedRange.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);

    await context.sync();

    if (usedRange.isNullObject) {
      addLog(`No existing output found on ${outputConfig.sheet}.`);
      return false;
    }

    const firstRowToClear = Math.max(usedRange.rowIndex, 1);
    const firstColumnToClear = Math.max(usedRange.columnIndex, 1);
    const lastRow = usedRange.rowIndex + usedRange.rowCount - 1;
    const lastColumn = usedRange.columnIndex + usedRange.columnCount - 1;
    const rowCount = lastRow - firstRowToClear + 1;
    const columnCount = lastColumn - firstColumnToClear + 1;

    if (rowCount <= 0 || columnCount <= 0) {
      addLog(`No clearable output found on ${outputConfig.sheet}.`);
      return false;
    }

    outputSheet
      .getRangeByIndexes(firstRowToClear, firstColumnToClear, rowCount, columnCount)
      .clear(Excel.ClearApplyTo.all);

    await context.sync();
    return true;
  });

  if (cleared) {
    addLog(`Cleared stale output from ${outputConfig.sheet}.`);
  }
}

async function writePayrollOutputs(outputConfig, outputs) {
  if (!outputs?.headcount?.table?.length) {
    throw new Error("Backend did not return a headcount output table.");
  }
  if (!outputs?.baseSalary?.total?.table?.length) {
    throw new Error("Backend did not return base salary output tables.");
  }
  if (!outputs?.bonusAccrual?.table?.length) {
    throw new Error("Backend did not return a bonus accrual output table.");
  }
  if (!outputs?.bonusPayout?.table?.length) {
    throw new Error("Backend did not return a bonus payout output table.");
  }
  if (!outputs?.severance?.table?.length) {
    throw new Error("Backend did not return a severance output table.");
  }

  await Excel.run(async (context) => {
    const outputSheet = context.workbook.worksheets.getItem(outputConfig.sheet);
    writeOutputTable(
      outputSheet,
      outputConfig.headcountStartCell,
      outputs.headcount.table,
      "0.00"
    );
    writeOutputTable(
      outputSheet,
      outputConfig.baseSalaryTotalStartCell,
      outputs.baseSalary.total.table,
      "#,##0"
    );
    writeOutputTable(
      outputSheet,
      outputConfig.baseSalaryDomesticStartCell,
      outputs.baseSalary.domestic.table,
      "#,##0"
    );
    writeOutputTable(
      outputSheet,
      outputConfig.baseSalaryInternationalStartCell,
      outputs.baseSalary.international.table,
      "#,##0"
    );
    writeOutputTable(
      outputSheet,
      outputConfig.baseSalaryCogsStartCell,
      outputs.baseSalary.cogs.table,
      "#,##0"
    );
    writeOutputTable(
      outputSheet,
      outputConfig.medicalStartCell,
      outputs.benefits.medical.table,
      "#,##0"
    );
    writeOutputTable(
      outputSheet,
      outputConfig.retirement401kStartCell,
      outputs.benefits.retirement401k.table,
      "#,##0"
    );
    writeOutputTable(
      outputSheet,
      outputConfig.otherBenefitsStartCell,
      outputs.benefits.otherBenefits.table,
      "#,##0"
    );
    writeOutputTable(
      outputSheet,
      outputConfig.bonusAccrualStartCell,
      outputs.bonusAccrual.table,
      "#,##0"
    );
    writeOutputTable(
      outputSheet,
      outputConfig.bonusPayoutStartCell,
      outputs.bonusPayout.table,
      "#,##0"
    );
    writeOutputTable(
      outputSheet,
      outputConfig.severanceStartCell,
      outputs.severance.table,
      "#,##0"
    );

    await context.sync();
  });

  addLog(
    `Payroll outputs written to ${outputConfig.sheet}.`
  );
}

function writeOutputTable(outputSheet, startCell, table, numberFormat) {
  const startRange = outputSheet.getRange(startCell);
  const targetRange = startRange.getResizedRange(
    table.length - 1,
    table[0].length - 1
  );

  targetRange.values = table;
  targetRange.numberFormat = buildOutputNumberFormat(table, numberFormat);
  targetRange.format.autofitColumns();
}

function buildOutputNumberFormat(table, numberFormat) {
  return table.map((row, rowIndex) =>
    row.map((_, columnIndex) => {
      if (rowIndex === 0 || columnIndex === 0) {
        return "@";
      }
      return numberFormat;
    })
  );
}

function updateMetrics(metrics) {
  const rowsPerSecond =
    metrics.loadTimeMs > 0 ? Math.round(metrics.totalRows / (metrics.loadTimeMs / 1000)) : 0;
  const percent =
    metrics.totalRows > 0
      ? ((metrics.includedRows / metrics.totalRows) * 100).toFixed(2)
      : "0.00";

  elements.includedRows.textContent = `${metrics.includedRows.toLocaleString()} of ${metrics.totalRows.toLocaleString()} (${percent}%)`;
  elements.loadSpeed.textContent = `${rowsPerSecond.toLocaleString()} rows/sec - ${(
    metrics.loadTimeMs / 1000
  ).toFixed(2)}s`;
  if (metrics.storeRows !== undefined) {
    addLog(
      `Detail storage selected for ${Number(metrics.storeRows).toLocaleString()} rows.`
    );
  }
}

function updateConfigUi(source, model) {
  elements.configState.textContent = "Config";
  elements.configState.className = "is-success";
  elements.dataSheet.textContent = source.sheet;
  elements.dataRange.textContent = source.dataRange;
  elements.filterRule.textContent = `${source.filterColumn} = 1`;
  elements.forecastPeriod.textContent = `${model.periods[0].label} to ${
    model.periods[model.periods.length - 1].label
  }`;
  elements.calculationMonths.textContent = `${model.calculationMonths} months`;
  addLog("Config loaded successfully.");
}

function setBusy(isBusy) {
  elements.button.disabled = isBusy;
  elements.button.textContent = isBusy ? "Loading..." : "Payroll Recalc";
}

function setStatus(message, type = "") {
  elements.status.textContent = message;
  elements.status.className = `status${type ? ` is-${type}` : ""}`;
}

async function checkBackendHealth() {
  const baseUrl = elements.backendUrl.value.trim().replace(/\/$/, "");
  if (!baseUrl) {
    setBackendUi("Missing", "error");
    return;
  }

  try {
    const response = await fetch(`${baseUrl}/health`);
    if (!response.ok) {
      throw new Error(`Health check returned ${response.status}`);
    }
    setBackendUi("Connected", "connected");
    elements.backendResult.textContent = "Ready";
    elements.backendResult.className = "is-success";
    addLog("Connection to backend verified.");
  } catch {
    setBackendUi("Offline", "error");
    elements.backendResult.textContent = "Offline";
    elements.backendResult.className = "is-error";
  }
}

function setBackendUi(label, state) {
  elements.backendState.textContent = label;
  elements.backendState.className = `state-pill is-${state}`;
  elements.cloudDot.className = `cloud-dot is-${state}`;
}

function addLog(message) {
  const item = document.createElement("li");
  const time = document.createElement("time");
  const text = document.createElement("span");
  time.textContent = new Date().toLocaleTimeString([], {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
  text.textContent = message;
  item.append(time, text);
  elements.activityLog.append(item);

  while (elements.activityLog.children.length > 8) {
    elements.activityLog.firstElementChild.remove();
  }
}

function safeHost(value) {
  try {
    return new URL(value).host;
  } catch {
    return "Backend";
  }
}
