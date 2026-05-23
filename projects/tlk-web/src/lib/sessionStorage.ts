export const SESSION_KEY_GITHUB_REPO = "tlkForgeGitHubRepo";
export const SESSION_KEY_GITHUB_BASE_BRANCH = "tlkForgeGitHubBaseBranch";
export const SESSION_KEY_GITHUB_CSV_FOLDER = "tlkForgeGitHubCsvFolder";
export const SESSION_KEY_GITHUB_TOKEN = "tlkForgeGitHubToken";

type SessionStorageLike = Pick<Storage, "getItem" | "setItem" | "removeItem"> | null;

function getSessionStorage(): SessionStorageLike {
  if (typeof window === "undefined") {
    return null;
  }
  try {
    return window.sessionStorage;
  } catch {
    return null;
  }
}

export function readSessionValue(
  key: string,
  fallback: string,
  storage: SessionStorageLike = getSessionStorage(),
): string {
  if (!storage) {
    return fallback;
  }
  try {
    const value = String(storage.getItem(key) || "").trim();
    return value || fallback;
  } catch {
    return fallback;
  }
}

export function writeSessionValue(
  key: string,
  value: string,
  storage: SessionStorageLike = getSessionStorage(),
): void {
  if (!storage) {
    return;
  }
  try {
    storage.setItem(key, String(value || "").trim());
  } catch {
    // ignore sessionStorage failures
  }
}

export function removeSessionValue(
  key: string,
  storage: SessionStorageLike = getSessionStorage(),
): void {
  if (!storage) {
    return;
  }
  try {
    storage.removeItem(key);
  } catch {
    // ignore sessionStorage failures
  }
}
