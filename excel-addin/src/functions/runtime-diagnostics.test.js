import assert from "node:assert/strict";
import test from "node:test";

import {
  batchQueueCheck,
  postCheck,
  runtimeCheck,
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

test("postCheck sends a POST request to backend debug endpoint", async () => {
  const requests = [];
  const result = await postCheck("https://example.test", {
    fetchFn: async (url, options) => {
      requests.push({ url, options });
      return { ok: true, status: 200 };
    },
  });

  assert.equal(result, "post=200");
  assert.equal(requests.length, 1);
  assert.equal(requests[0].url, "https://example.test/debug/client-log");
  assert.equal(requests[0].options.method, "POST");
  assert.equal(JSON.parse(requests[0].options.body).source, "HCA.POST_CHECK");
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
