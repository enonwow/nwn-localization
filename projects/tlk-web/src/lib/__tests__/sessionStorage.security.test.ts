import { describe, expect, it } from "vitest";

import {
  readSessionValue,
  removeSessionValue,
  SESSION_KEY_GITHUB_BASE_BRANCH,
  SESSION_KEY_GITHUB_CSV_FOLDER,
  SESSION_KEY_GITHUB_REPO,
  SESSION_KEY_GITHUB_TOKEN,
  writeSessionValue,
} from "../sessionStorage";

type MemoryStorageOptions = {
  throwOnGet?: boolean;
  throwOnSet?: boolean;
  throwOnRemove?: boolean;
};

function createMemoryStorage(
  seed: Record<string, string> = {},
  options: MemoryStorageOptions = {},
): Pick<Storage, "getItem" | "setItem" | "removeItem"> {
  const map = new Map(Object.entries(seed));
  return {
    getItem(key: string) {
      if (options.throwOnGet) {
        throw new Error("getItem failed");
      }
      return map.has(key) ? map.get(key)! : null;
    },
    setItem(key: string, value: string) {
      if (options.throwOnSet) {
        throw new Error("setItem failed");
      }
      map.set(key, String(value));
    },
    removeItem(key: string) {
      if (options.throwOnRemove) {
        throw new Error("removeItem failed");
      }
      map.delete(key);
    },
  };
}

describe("sessionStorage helpers", () => {
  it("reads trimmed values and falls back for empty key", () => {
    const storage = createMemoryStorage({
      repo: "  enonwow/nwn-localization-test  ",
      blank: "   ",
    });

    expect(readSessionValue("repo", "fallback", storage)).toBe("enonwow/nwn-localization-test");
    expect(readSessionValue("blank", "fallback", storage)).toBe("fallback");
    expect(readSessionValue("missing", "fallback", storage)).toBe("fallback");
  });

  it("uses fallback when storage is unavailable or throws", () => {
    const throwingStorage = createMemoryStorage({}, { throwOnGet: true });

    expect(readSessionValue("repo", "fallback", null)).toBe("fallback");
    expect(readSessionValue("repo", "fallback", throwingStorage)).toBe("fallback");
  });

  it("writes trimmed values and swallows storage set errors", () => {
    const storage = createMemoryStorage();
    writeSessionValue("repo", "  enonwow/nwn-localization-test  ", storage);
    expect(readSessionValue("repo", "fallback", storage)).toBe("enonwow/nwn-localization-test");

    const throwingStorage = createMemoryStorage({}, { throwOnSet: true });
    expect(() => writeSessionValue("repo", "value", throwingStorage)).not.toThrow();
  });

  it("removes keys and swallows remove errors", () => {
    const storage = createMemoryStorage({ repo: "enonwow/nwn-localization-test" });
    removeSessionValue("repo", storage);
    expect(readSessionValue("repo", "fallback", storage)).toBe("fallback");

    const throwingStorage = createMemoryStorage({ repo: "x" }, { throwOnRemove: true });
    expect(() => removeSessionValue("repo", throwingStorage)).not.toThrow();
  });

  it("persists repo settings but never keeps PAT in session storage", () => {
    const storage = createMemoryStorage({
      [SESSION_KEY_GITHUB_TOKEN]: "secret-token",
    });

    removeSessionValue(SESSION_KEY_GITHUB_TOKEN, storage);
    writeSessionValue(SESSION_KEY_GITHUB_REPO, "enonwow/nwn-localization-test", storage);
    writeSessionValue(SESSION_KEY_GITHUB_BASE_BRANCH, "main", storage);
    writeSessionValue(SESSION_KEY_GITHUB_CSV_FOLDER, "csv-latest", storage);

    expect(readSessionValue(SESSION_KEY_GITHUB_TOKEN, "", storage)).toBe("");
    expect(readSessionValue(SESSION_KEY_GITHUB_REPO, "", storage)).toBe("enonwow/nwn-localization-test");
    expect(readSessionValue(SESSION_KEY_GITHUB_BASE_BRANCH, "", storage)).toBe("main");
    expect(readSessionValue(SESSION_KEY_GITHUB_CSV_FOLDER, "", storage)).toBe("csv-latest");
  });
});
