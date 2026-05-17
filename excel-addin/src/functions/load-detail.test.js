import assert from "node:assert/strict";
import test from "node:test";

import {
  buildClientLogContext,
  buildLoadDetailUrl,
  buildLoadDetailLookupKey,
  getBackendHealthStatus,
  loadDetail,
  normalizePeriodEndDate,
  parseLoadDetailValue,
  queueLoadDetailLookup,
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

test("buildLoadDetailLookupKey normalizes lookup fields", () => {
  assert.equal(
    buildLoadDetailLookupKey(" User@Example.COM ", {
      outputKey: " payroll.output.401k ",
      periodEndDate: " 2026-05-31 ",
      unitId: " E1 ",
    }),
    "user@example.com\u001fpayroll.output.401k\u001f2026-05-31\u001fE1"
  );
});

test("queueLoadDetailLookup batches 1000 lookups into 2 requests", async () => {
  const requests = [];
  const timers = [];

  async function fetchFn(url, options) {
    const body = JSON.parse(options.body);
    requests.push({ url, body });
    return {
      ok: true,
      status: 200,
      json: async () => ({
        status: "ok",
        values: body.items.map((_, index) => index + 1),
        foundCount: body.items.length,
      }),
    };
  }

  const promises = [];
  for (let index = 0; index < 1000; index += 1) {
    promises.push(
      queueLoadDetailLookup(
        "https://example.test",
        "user@example.com",
        {
          outputKey: "payroll.output.base_salary_total",
          periodEndDate: "2026-05-31",
          unitId: `E${index}`,
        },
        {
          fetchFn,
          maxBatchSize: 500,
          setTimeoutFn: (callback) => {
            timers.push(callback);
            return timers.length;
          },
        }
      )
    );
  }

  assert.equal(timers.length, 1);
  timers[0]();
  const values = await Promise.all(promises);

  assert.equal(requests.length, 2);
  assert.equal(requests[0].url, "https://example.test/payroll/load-detail-batch");
  assert.equal(requests[0].body.items.length, 500);
  assert.equal(requests[1].body.items.length, 500);
  assert.equal(values.length, 1000);
  assert.equal(values[0], 1);
  assert.equal(values[499], 500);
  assert.equal(values[500], 1);
  assert.equal(values[999], 500);
});

test("queueLoadDetailLookup dedupes identical lookups in one pending batch", async () => {
  const requests = [];
  const timers = [];

  async function fetchFn(url, options) {
    const body = JSON.parse(options.body);
    requests.push({ url, body });
    return {
      ok: true,
      status: 200,
      json: async () => ({ status: "ok", values: [432], foundCount: 1 }),
    };
  }

  const promises = [];
  for (let index = 0; index < 10; index += 1) {
    promises.push(
      queueLoadDetailLookup(
        "https://example.test",
        "user@example.com",
        {
          outputKey: "payroll.output.401k",
          periodEndDate: "2026-05-31",
          unitId: "E1",
        },
        {
          fetchFn,
          setTimeoutFn: (callback) => {
            timers.push(callback);
            return timers.length;
          },
        }
      )
    );
  }

  timers[0]();
  const values = await Promise.all(promises);

  assert.deepEqual(values, Array(10).fill(432));
  assert.equal(requests.length, 1);
  assert.equal(requests[0].body.items.length, 1);
});

test("queueLoadDetailLookup keeps different backend URLs in separate batches", async () => {
  const requests = [];
  const timers = [];

  async function fetchFn(url, options) {
    const body = JSON.parse(options.body);
    requests.push({ url, body });
    const value = url.includes("first.example.test") ? 1 : 2;
    return {
      ok: true,
      status: 200,
      json: async () => ({ status: "ok", values: [value], foundCount: 1 }),
    };
  }

  const options = {
    fetchFn,
    setTimeoutFn: (callback) => {
      timers.push(callback);
      return timers.length;
    },
  };

  const first = queueLoadDetailLookup(
    "https://first.example.test",
    "user@example.com",
    {
      outputKey: "payroll.output.401k",
      periodEndDate: "2026-05-31",
      unitId: "E1",
    },
    options
  );
  const second = queueLoadDetailLookup(
    "https://second.example.test",
    "user@example.com",
    {
      outputKey: "payroll.output.401k",
      periodEndDate: "2026-05-31",
      unitId: "E1",
    },
    options
  );

  assert.equal(timers.length, 2);
  timers.forEach((callback) => callback());

  assert.deepEqual(await Promise.all([first, second]), [1, 2]);
  assert.equal(requests.length, 2);
  assert.equal(
    requests[0].url,
    "https://first.example.test/payroll/load-detail-batch"
  );
  assert.equal(
    requests[1].url,
    "https://second.example.test/payroll/load-detail-batch"
  );
});

test("queueLoadDetailLookup rejects when response length does not match", async () => {
  const timers = [];
  const promise = queueLoadDetailLookup(
    "https://example.test",
    "user@example.com",
    {
      outputKey: "payroll.output.401k",
      periodEndDate: "2026-05-31",
      unitId: "E1",
    },
    {
      fetchFn: async () => ({
        ok: true,
        status: 200,
        json: async () => ({ status: "ok", values: [], foundCount: 0 }),
      }),
      setTimeoutFn: (callback) => {
        timers.push(callback);
        return timers.length;
      },
    }
  );

  timers[0]();

  await assert.rejects(
    promise,
    /LOAD_DETAIL batch returned 0 values for 1 lookups/
  );
});

test("queueLoadDetailLookup works when global setTimeout is unavailable", async () => {
  const previousSetTimeout = globalThis.setTimeout;
  const requests = [];
  globalThis.setTimeout = undefined;

  try {
    const value = await queueLoadDetailLookup(
      "https://example.test",
      "user@example.com",
      {
        outputKey: "payroll.output.401k",
        periodEndDate: "2026-05-31",
        unitId: "E1",
      },
      {
        fetchFn: async (url, options) => {
          requests.push({ url, body: JSON.parse(options.body) });
          return {
            ok: true,
            status: 200,
            json: async () => ({ status: "ok", values: [432], foundCount: 1 }),
          };
        },
      }
    );

    assert.equal(value, 432);
    assert.equal(requests.length, 1);
    assert.equal(requests[0].url, "https://example.test/payroll/load-detail-batch");
  } finally {
    globalThis.setTimeout = previousSetTimeout;
  }
});

test("loadDetail uses optional user key override for single lookup request", async () => {
  const requests = [];
  const previousFetch = globalThis.fetch;
  const previousLocalStorage = globalThis.localStorage;

  globalThis.localStorage = {
    getItem: (key) => (key === "xf1.backendUrl" ? "https://example.test" : ""),
    setItem: () => {},
  };
  globalThis.fetch = async (url, options) => {
    requests.push({ url, options });
    return {
      ok: true,
      status: 200,
      json: async () => ({ status: "found", value: 99 }),
    };
  };

  try {
    const value = await loadDetail(
      "payroll.output.401k",
      "2026-05-15",
      "E1",
      "override@example.com"
    );

    assert.equal(value, 99);
    assert.equal(requests.length, 1);
    assert.equal(
      requests[0].url,
      "https://example.test/payroll/load-detail?userKey=override%40example.com&outputKey=payroll.output.401k&periodEndDate=2026-05-31&unitId=E1"
    );
    assert.equal(requests[0].options, undefined);
  } finally {
    globalThis.fetch = previousFetch;
    globalThis.localStorage = previousLocalStorage;
  }
});

test("production functions metadata exposes LOAD_DETAIL and temporary runtime checks", async () => {
  const metadata = JSON.parse(
    await import("node:fs/promises").then((fs) =>
      fs.readFile(new URL("../../public/functions.json", import.meta.url), "utf8")
    )
  );

  assert.deepEqual(
    metadata.functions.map((fn) => fn.id),
    [
      "LOAD_DETAIL",
      "RUNTIME_CHECK",
      "POST_CHECK",
      "BATCH_QUEUE_CHECK",
      "TEXT_BATCH_CHECK",
    ]
  );
});
