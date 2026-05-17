import assert from "node:assert/strict";
import test from "node:test";

import {
  batchQueueCheck,
  postCheck,
  runtimeCheck,
  textBatchCheck,
} from "./runtime-diagnostics.js";

test("runtimeCheck reports promise, timer, fetch, and storage support", async () => {
  const result = await runtimeCheck({
    setTimeoutFn: (callback) => {
      callback();
      return 1;
    },
    fetchFn: async () => {},
    storage: {},
  });

  assert.match(result, /promise=ok/);
  assert.match(result, /setTimeout=function/);
  assert.match(result, /timer=ok/);
  assert.match(result, /fetch=function/);
  assert.match(result, /officeStorage=available/);
});

test("runtimeCheck reports missing timer without failing", async () => {
  const result = await runtimeCheck({
    setTimeoutFn: undefined,
    fetchFn: async () => {},
    storage: null,
  });

  assert.match(result, /setTimeout=missing/);
  assert.match(result, /timer=missing/);
});

test("postCheck sends a JSON POST request to backend debug endpoint", async () => {
  const requests = [];
  const result = await postCheck("json", "https://example.test", {
    fetchFn: async (url, options) => {
      requests.push({ url, options });
      return { ok: true, status: 200 };
    },
  });

  assert.equal(result, "mode=json;post=200");
  assert.equal(requests.length, 1);
  assert.equal(requests[0].url, "https://example.test/debug/client-log");
  assert.equal(requests[0].options.method, "POST");
  assert.equal(requests[0].options.headers["Content-Type"], "application/json");
  assert.equal(JSON.parse(requests[0].options.body).source, "HCA.POST_CHECK");
});

test("postCheck can send text/plain POST without JSON content type", async () => {
  const requests = [];
  const result = await postCheck("text", "https://example.test", {
    fetchFn: async (url, options) => {
      requests.push({ url, options });
      return { ok: false, status: 422 };
    },
  });

  assert.equal(result, "mode=text;post=422");
  assert.equal(requests[0].options.headers["Content-Type"], "text/plain");
  assert.match(requests[0].options.body, /runtime=/);
});

test("postCheck can send an empty POST with no headers or body", async () => {
  const requests = [];
  const result = await postCheck("empty", "https://example.test", {
    fetchFn: async (url, options) => {
      requests.push({ url, options });
      return { ok: false, status: 422 };
    },
  });

  assert.equal(result, "mode=empty;post=422");
  assert.deepEqual(requests[0].options, { method: "POST" });
});

test("postCheck can send form-urlencoded POST", async () => {
  const requests = [];
  const result = await postCheck("form", "https://example.test", {
    fetchFn: async (url, options) => {
      requests.push({ url, options });
      return { ok: false, status: 422 };
    },
  });

  assert.equal(result, "mode=form;post=422");
  assert.equal(
    requests[0].options.headers["Content-Type"],
    "application/x-www-form-urlencoded"
  );
  assert.match(requests[0].options.body, /source=HCA.POST_CHECK/);
});

test("postCheck treats the first argument as backend URL for old one-argument calls", async () => {
  const requests = [];
  await postCheck("https://example.test", {
    fetchFn: async (url, options) => {
      requests.push({ url, options });
      return { ok: true, status: 200 };
    },
  });

  assert.equal(requests[0].url, "https://example.test/debug/client-log");
  assert.equal(requests[0].options.headers["Content-Type"], "application/json");
});

test("batchQueueCheck groups calls sharing a token before timer flush", async () => {
  const timers = [];
  const options = {
    setTimeoutFn: (callback) => {
      timers.push(callback);
      return timers.length;
    },
  };

  const first = batchQueueCheck("same-token", 50, options);
  const second = batchQueueCheck("same-token", 50, options);

  assert.equal(timers.length, 1);
  timers[0]();

  const values = await Promise.all([first, second]);
  assert.match(values[0], /token=same-token/);
  assert.match(values[0], /index=1/);
  assert.match(values[0], /size=2/);
  assert.match(values[1], /index=2/);
  assert.match(values[1], /size=2/);
});

test("textBatchCheck sends two lookups as text/plain batch payload", async () => {
  const requests = [];
  const result = await textBatchCheck(
    "user@example.com",
    "payroll.output.base_salary_total",
    "2026-05-15",
    "E1",
    "text",
    "https://example.test",
    {
      fetchFn: async (url, options) => {
        requests.push({ url, options });
        return {
          ok: true,
          status: 200,
          json: async () => ({ status: "ok", values: [99, 99], foundCount: 2 }),
        };
      },
    }
  );

  assert.equal(result, "mode=text;status=ok;count=2;first=99");
  assert.equal(requests[0].url, "https://example.test/payroll/load-detail-batch-text");
  assert.equal(requests[0].options.headers["Content-Type"], "text/plain");
  assert.equal(JSON.parse(requests[0].options.body).items.length, 2);
});

test("textBatchCheck can send form encoded batch payload", async () => {
  const requests = [];
  const result = await textBatchCheck(
    "user@example.com",
    "payroll.output.base_salary_total",
    "2026-05-15",
    "E1",
    "form",
    "https://example.test",
    {
      fetchFn: async (url, options) => {
        requests.push({ url, options });
        return {
          ok: true,
          status: 200,
          json: async () => ({ status: "ok", values: [88, 88], foundCount: 2 }),
        };
      },
    }
  );

  assert.equal(result, "mode=form;status=ok;count=2;first=88");
  assert.equal(
    requests[0].options.headers["Content-Type"],
    "application/x-www-form-urlencoded"
  );
  assert.match(requests[0].options.body, /^payload=/);
});
