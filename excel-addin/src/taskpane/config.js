const CONFIG_NAME = "HCA.Engine.Config";

export function getConfigNamedRange(context) {
  return context.workbook.names.getItem(CONFIG_NAME).getRange();
}

export function parseConfig(values) {
  if (!values || values.length < 2) {
    throw new Error(`Named range ${CONFIG_NAME} does not contain the expected config table.`);
  }

  const headerRow = values[0].map((value) => normalizeKey(value));
  const columns = {
    key: headerRow.indexOf("key"),
    value: headerRow.indexOf("value"),
  };

  for (const [name, index] of Object.entries(columns)) {
    if (index === -1) {
      throw new Error(`Config named range is missing required column: ${name}`);
    }
  }

  const settings = {};
  values.slice(1).forEach((row) => {
    const key = normalizeKey(row[columns.key]);
    if (!key) {
      return;
    }
    settings[key] = row[columns.value];
  });

  const lastActualsDate = parseExcelDate(
    requiredSetting(settings, "model.last_actuals_date"),
    "model.last_actuals_date"
  );
  const modelEndDate = parseExcelDate(
    requiredSetting(settings, "model.model_end_date"),
    "model.model_end_date"
  );
  const financialYearEndMonth = parseMonthNumber(
    requiredSetting(settings, "model.financial_year_end_month"),
    "model.financial_year_end_month"
  );
  const timeline = buildModelTimeline(
    lastActualsDate,
    modelEndDate,
    financialYearEndMonth
  );

  const dataReference = parseSheetReference(
    requiredSetting(settings, "payroll.data_range"),
    "payroll.data_range"
  );
  const headersReference = parseSheetReference(
    requiredSetting(settings, "payroll.headers_range"),
    "payroll.headers_range"
  );
  const filterReference = parseSheetReference(
    requiredSetting(settings, "payroll.filter_column"),
    "payroll.filter_column"
  );
  const storeFilterReference = parseSheetReference(
    requiredSetting(settings, "payroll.store_filter_column"),
    "payroll.store_filter_column"
  );

  if (headersReference.sheet !== dataReference.sheet) {
    throw new Error("Payroll headers and data ranges must be on the same sheet.");
  }
  if (filterReference.sheet !== dataReference.sheet) {
    throw new Error("Payroll filter column must be on the same sheet as the data range.");
  }
  if (storeFilterReference.sheet !== dataReference.sheet) {
    throw new Error("Payroll store filter column must be on the same sheet as the data range.");
  }

  const output = buildOutputConfig(settings);
  const assumptions = {
    benefits: {
      medical: {
        domestic: parseNumberSetting(
          requiredSetting(settings, "payroll.benefits.medical.domestic"),
          "payroll.benefits.medical.domestic"
        ),
        international: parseNumberSetting(
          requiredSetting(settings, "payroll.benefits.medical.international"),
          "payroll.benefits.medical.international"
        ),
      },
      retirement401k: {
        domestic: parseNumberSetting(
          requiredSetting(settings, "payroll.benefits.401k.domestic"),
          "payroll.benefits.401k.domestic"
        ),
        international: parseNumberSetting(
          requiredSetting(settings, "payroll.benefits.401k.international"),
          "payroll.benefits.401k.international"
        ),
      },
      otherBenefits: {
        domestic: parseNumberSetting(
          requiredSetting(settings, "payroll.benefits.other.domestic"),
          "payroll.benefits.other.domestic"
        ),
        international: parseNumberSetting(
          requiredSetting(settings, "payroll.benefits.other.international"),
          "payroll.benefits.other.international"
        ),
      },
    },
    bonus: {
      cap: parseNumberSetting(
        requiredSetting(settings, "payroll.bonus_cap"),
        "payroll.bonus_cap"
      ),
      executivePlan: {
        netNewArrWeight: parseNumberSetting(
          requiredSetting(settings, "payroll.exec_bonus_nnar_weight"),
          "payroll.exec_bonus_NNAR_weight"
        ),
        burnMultipleWeight: parseNumberSetting(
          requiredSetting(settings, "payroll.exec_bonus_burn_multiple_weight"),
          "payroll.exec_bonus_burn_multiple_weight"
        ),
      },
      incentivePlan: {
        netNewArrWeight: parseNumberSetting(
          requiredSetting(settings, "payroll.incentive_bonus_nnar_weight"),
          "payroll.incentive_bonus_NNAR_weight"
        ),
        burnMultipleWeight: parseNumberSetting(
          requiredSetting(settings, "payroll.incentive_bonus_burn_multiple_weight"),
          "payroll.incentive_bonus_burn_multiple_weight"
        ),
      },
    },
  };
  const seriesRanges = {
    netNewArrAchieved: parseSheetReference(
      requiredSetting(settings, "payroll.net_new_arr_achieved"),
      "payroll.net_new_ARR_achieved"
    ),
    burnMultipleAchieved: parseSheetReference(
      requiredSetting(settings, "payroll.burn_multiple_achieved"),
      "payroll.burn_multiple_achieved"
    ),
  };

  return {
    model: {
      lastActualsDate: formatIsoDate(lastActualsDate),
      modelEndDate: formatIsoDate(modelEndDate),
      calculationStartDate: timeline.calculationStartDate,
      calculationEndDate: timeline.calculationEndDate,
      calculationMonths: timeline.periods.length,
      financialYearEndMonth,
      periods: timeline.periods,
    },
    payroll: {
      dataLoadSheet: dataReference.sheet,
      cellRange: dataReference.address,
      headers: headersReference.address,
      filterColumn: extractStartColumn(filterReference.address),
      storeFilterColumn: extractStartColumn(storeFilterReference.address),
    },
    output,
    assumptions,
    seriesRanges,
  };
}

function buildOutputConfig(settings) {
  const references = {
    headcountStartCell: parseSheetReference(
      requiredSetting(settings, "payroll.output.headcount"),
      "payroll.output.headcount"
    ),
    baseSalaryTotalStartCell: parseSheetReference(
      requiredSetting(settings, "payroll.output.base_salary_total"),
      "payroll.output.base_salary_total"
    ),
    baseSalaryDomesticStartCell: parseSheetReference(
      requiredSetting(settings, "payroll.output.base_salary_domestic"),
      "payroll.output.base_salary_domestic"
    ),
    baseSalaryInternationalStartCell: parseSheetReference(
      requiredSetting(settings, "payroll.output.base_salary_international"),
      "payroll.output.base_salary_international"
    ),
    baseSalaryCogsStartCell: parseSheetReference(
      requiredSetting(settings, "payroll.output.base_salary_cogs"),
      "payroll.output.base_salary_cogs"
    ),
    medicalStartCell: parseSheetReference(
      requiredSetting(settings, "payroll.output.medical"),
      "payroll.output.medical"
    ),
    retirement401kStartCell: parseSheetReference(
      requiredSetting(settings, "payroll.output.401k"),
      "payroll.output.401k"
    ),
    otherBenefitsStartCell: parseSheetReference(
      requiredSetting(settings, "payroll.output.other_benefits"),
      "payroll.output.other_benefits"
    ),
    bonusAccrualStartCell: parseSheetReference(
      requiredSetting(settings, "payroll.output.bonus_accrual"),
      "payroll.output.bonus_accrual"
    ),
    bonusPayoutStartCell: parseSheetReference(
      requiredSetting(settings, "payroll.output.bonus_payout"),
      "payroll.output.bonus_payout"
    ),
    severanceStartCell: parseSheetReference(
      requiredSetting(settings, "payroll.output.severance"),
      "payroll.output.severance"
    ),
  };
  const sheet = references.headcountStartCell.sheet;

  Object.entries(references).forEach(([label, reference]) => {
    if (reference.sheet !== sheet) {
      throw new Error(`${label} must be on the same output sheet as headcount.`);
    }
  });

  return Object.entries(references).reduce(
    (output, [key, reference]) => {
      output[key] = reference.address;
      return output;
    },
    { sheet }
  );
}

function parseSheetReference(value, label) {
  const reference = String(value ?? "").trim();
  const bangIndex = reference.lastIndexOf("!");

  if (bangIndex <= 0 || bangIndex === reference.length - 1) {
    throw new Error(`${label} must be a sheet-qualified reference, for example PayrollData!B5:R1531.`);
  }

  return {
    sheet: reference.slice(0, bangIndex).replace(/^'|'$/g, ""),
    address: reference.slice(bangIndex + 1).replace(/\$/g, ""),
  };
}

export function getFilterOffset(rangeAddress, filterColumn) {
  const startColumn = extractStartColumn(rangeAddress);
  const startIndex = columnToNumber(startColumn);
  const filterIndex = columnToNumber(filterColumn);
  const offset = filterIndex - startIndex;
  const columnCount = getRangeColumnCount(rangeAddress);

  if (offset < 0) {
    throw new Error(
      `Filter column ${filterColumn} is outside data range ${rangeAddress}.`
    );
  }

  if (offset >= columnCount) {
    throw new Error(
      `Filter column ${filterColumn} is outside data range ${rangeAddress}. Expand payroll.data_range and payroll.headers_range to include that column.`
    );
  }

  return offset;
}

export function isSelectedFlag(value) {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "number") {
    return value === 1;
  }

  const text = String(value ?? "").trim().toLowerCase();
  return text === "1" || text === "true" || text === "yes";
}

function getRangeColumnCount(rangeAddress) {
  const cleaned = String(rangeAddress).split("!").pop().replace(/\$/g, "");
  const [startAddress, endAddress = startAddress] = cleaned.split(":");
  const startColumn = extractStartColumn(startAddress);
  const endColumn = extractStartColumn(endAddress);
  return columnToNumber(endColumn) - columnToNumber(startColumn) + 1;
}

function extractStartColumn(rangeAddress) {
  const cleaned = String(rangeAddress).split("!").pop().replace(/\$/g, "");
  const match = cleaned.match(/^([A-Z]+)(?:\d+)?/i);

  if (!match) {
    throw new Error(`Could not read start column from range: ${rangeAddress}`);
  }

  return match[1];
}

function columnToNumber(columnLetters) {
  return String(columnLetters)
    .trim()
    .toUpperCase()
    .split("")
    .reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);
}

function requiredSetting(settings, key) {
  const value = settings[key];
  if (value === undefined || value === null || String(value).trim() === "") {
    throw new Error(`Config value is blank or missing: ${key}`);
  }
  return value;
}

function normalizeKey(value) {
  return String(value ?? "").trim().toLowerCase();
}

function parseExcelDate(value, label) {
  if (typeof value === "number") {
    return new Date(Date.UTC(1899, 11, 30 + value));
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
  }

  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    throw new Error(`Config date is invalid: ${label}`);
  }

  return new Date(
    Date.UTC(parsed.getFullYear(), parsed.getMonth(), parsed.getDate())
  );
}

function parseMonthNumber(value, label) {
  const month = Number(value);
  if (!Number.isInteger(month) || month < 1 || month > 12) {
    throw new Error(`${label} must be a number from 1 to 12.`);
  }
  return month;
}

function parseNumberSetting(value, label) {
  if (typeof value === "number") {
    return value;
  }

  const parsed = Number(String(value).replace(/,/g, "").trim());
  if (Number.isNaN(parsed)) {
    throw new Error(`${label} must be a number.`);
  }

  return parsed;
}

function buildModelTimeline(lastActualsDate, modelEndDate, financialYearEndMonth) {
  const startDate = endOfMonth(addMonths(lastActualsDate, 1));
  const endDate = endOfMonth(modelEndDate);

  if (startDate > endDate) {
    throw new Error("Model end date must be after Last actuals date.");
  }

  const periods = [];
  for (
    let cursor = startDate;
    cursor <= endDate;
    cursor = endOfMonth(addMonths(cursor, 1))
  ) {
    periods.push({
      date: formatIsoDate(cursor),
      label: formatMonthLabel(cursor),
      financialYear: getFinancialYear(cursor, financialYearEndMonth),
    });
  }

  return {
    calculationStartDate: formatIsoDate(startDate),
    calculationEndDate: formatIsoDate(endDate),
    periods,
  };
}

function addMonths(date, months) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + months, 1));
}

function endOfMonth(date) {
  return new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth() + 1, 0));
}

function getFinancialYear(date, financialYearEndMonth) {
  const month = date.getUTCMonth() + 1;
  const year = date.getUTCFullYear();
  return month <= financialYearEndMonth ? year : year + 1;
}

function formatIsoDate(date) {
  return date.toISOString().slice(0, 10);
}

function formatMonthLabel(date) {
  return date.toLocaleString("en-US", {
    month: "short",
    year: "numeric",
    timeZone: "UTC",
  });
}
