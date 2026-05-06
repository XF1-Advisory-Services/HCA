export const USER_KEY_STORAGE_KEY = "xf1.userKey";
export const BACKEND_URL_STORAGE_KEY = "xf1.backendUrl";

export function normalizeUserKey(value) {
  return String(value ?? "").trim().toLowerCase();
}

export async function readSharedSetting(key) {
  if (globalThis.OfficeRuntime?.storage) {
    const value = await globalThis.OfficeRuntime.storage.getItem(key);
    if (value) {
      return value;
    }
  }

  return globalThis.localStorage?.getItem(key) || "";
}

export async function writeSharedSetting(key, value) {
  const normalizedValue = String(value ?? "");
  globalThis.localStorage?.setItem(key, normalizedValue);

  if (globalThis.OfficeRuntime?.storage) {
    await globalThis.OfficeRuntime.storage.setItem(key, normalizedValue);
  }
}
