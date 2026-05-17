import { loadDetail } from "./load-detail.js";
import {
  batchQueueCheck,
  postCheck,
  runtimeCheck,
} from "./runtime-diagnostics.js";

if (globalThis.CustomFunctions?.associate) {
  globalThis.CustomFunctions.associate("LOAD_DETAIL", loadDetail);
  globalThis.CustomFunctions.associate("RUNTIME_CHECK", runtimeCheck);
  globalThis.CustomFunctions.associate("POST_CHECK", postCheck);
  globalThis.CustomFunctions.associate("BATCH_QUEUE_CHECK", batchQueueCheck);
}
