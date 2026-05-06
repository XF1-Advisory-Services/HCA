import assert from "node:assert/strict";
import test from "node:test";

import { normalizePeriodEndDate, parseLoadDetailValue } from "./load-detail.js";

test("normalizePeriodEndDate accepts an Excel date serial", () => {
  assert.equal(normalizePeriodEndDate(46263), "2026-08-31");
});

test("normalizePeriodEndDate accepts an ISO date string and returns month end", () => {
  assert.equal(normalizePeriodEndDate("2026-05-15"), "2026-05-31");
});

test("parseLoadDetailValue returns zero when backend has no stored row", () => {
  assert.equal(parseLoadDetailValue({ value: null }), 0);
});
