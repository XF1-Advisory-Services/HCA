import assert from "node:assert/strict";
import test from "node:test";

import {
  buildClientLogContext,
  buildLoadDetailUrl,
  diag,
  diagLoadDetail,
  getBackendHealthStatus,
  normalizePeriodEndDate,
  parseLoadDetailValue,
} from "./load-detail.js";

test("normalizePeriodEndDate accepts an Excel date serial", () => {
  assert.equal(normalizePeriodEndDate(46263), "2026-08-31");
});

test("normalizePeriodEndDate accepts an ISO date string and returns month end", () => {
  assert.equal(normalizePeriodEndDate("2026-05-15"), "2026-05-31");
});

test("parseLoadDetailValue returns zero when backend has no stored row", () => {
  assert.equal(parseLoadDetailValue({ value: null }), 0);
});

test("buildClientLogContext captures diagnostic fields without payroll values", () => {
  assert.deepEqual(
    buildClientLogContext(
      " payroll.output.401k ",
      "30/Apr/26",
      " EX18 ",
      "user@example.com"
    ),
    {
      outputKey: "payroll.output.401k",
      periodType: "string",
      periodRaw: "30/Apr/26",
      unitId: "EX18",
      userKeyOverrideProvided: true,
    }
  );
});

test("diag returns a plain value without backend access", () => {
  assert.equal(diag(), 123);
});

test("getBackendHealthStatus returns response status when backend is reachable", async () => {
  const status = await getBackendHealthStatus("https://example.test", async (url) => {
    assert.equal(url, "https://example.test/health");
    return { status: 200 };
  });

  assert.equal(status, 200);
});

test("getBackendHealthStatus returns -1 when fetch fails", async () => {
  const status = await getBackendHealthStatus("https://example.test", async () => {
    throw new Error("Failed to fetch");
  });

  assert.equal(status, -1);
});

test("diagLoadDetail posts known-good lookup values", async () => {
  const value = await diagLoadDetail(async (url, options) => {
    assert.equal(
      url,
      "https://hca-calc-engine.onrender.com/payroll/load-detail?userKey=vavrinec%40xf1advisory.com&outputKey=payroll.output.401k&periodEndDate=2026-04-30&unitId=EX18"
    );
    assert.equal(options, undefined);
    return {
      ok: true,
      status: 200,
      json: async () => ({ status: "found", value: 432.365 }),
    };
  });

  assert.equal(value, 432.365);
});

test("diagLoadDetail returns -1 when known-good lookup request fails", async () => {
  const value = await diagLoadDetail(async () => {
    throw new Error("Failed to fetch");
  });

  assert.equal(value, -1);
});

test("buildLoadDetailUrl encodes query parameters", () => {
  assert.equal(
    buildLoadDetailUrl("https://example.test/", {
      userKey: "user@example.com",
      outputKey: "payroll.output.base_salary_total",
      periodEndDate: "2026-05-31",
      unitId: "EX 18",
    }),
    "https://example.test/payroll/load-detail?userKey=user%40example.com&outputKey=payroll.output.base_salary_total&periodEndDate=2026-05-31&unitId=EX+18"
  );
});
