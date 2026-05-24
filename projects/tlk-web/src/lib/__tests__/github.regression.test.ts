import { describe, expect, it, vi } from "vitest";

import { fetchRepoFileFromGitHub, publishCsvToGitHub } from "../github";

describe("GitHub regressions", () => {
  it("keeps plain-text API error details when JSON parsing fails", async () => {
    const originalFetch = globalThis.fetch;
    const fetchMock = vi.fn(async () => ({
      ok: false,
      status: 500,
      json: async () => {
        throw new Error("invalid json");
      },
      text: async () => "fatal backend error",
    }));

    (globalThis as { fetch: typeof fetch }).fetch = fetchMock as unknown as typeof fetch;
    try {
      await expect(
        fetchRepoFileFromGitHub({
          repoFullName: "enonwow/nwn-localization-test",
          branch: "main",
          filePath: "csv-latest/tlc.csv",
        }),
      ).rejects.toThrow("GitHub API 500: fatal backend error");
    } finally {
      (globalThis as { fetch: typeof fetch }).fetch = originalFetch;
    }
  });

  it("rejects empty repository file payload", async () => {
    const originalFetch = globalThis.fetch;
    const fetchMock = vi.fn(async () => ({
      ok: true,
      status: 200,
      arrayBuffer: async () => new ArrayBuffer(0),
    }));

    (globalThis as { fetch: typeof fetch }).fetch = fetchMock as unknown as typeof fetch;
    try {
      await expect(
        fetchRepoFileFromGitHub({
          repoFullName: "enonwow/nwn-localization-test",
          branch: "main",
          filePath: "csv-latest/tlc.csv",
        }),
      ).rejects.toThrow("Repository file is empty");
    } finally {
      (globalThis as { fetch: typeof fetch }).fetch = originalFetch;
    }
  });

  it("rejects publish when normalized CSV file name is empty", async () => {
    await expect(
      publishCsvToGitHub({
        token: "test-token",
        repoFullName: "enonwow/nwn-localization-test",
        baseBranch: "main",
        csvFolder: "csv-latest",
        fileName: "///",
        csvBytes: new TextEncoder().encode("A,B\r\n1,2"),
      }),
    ).rejects.toThrow("CSV file name is missing");
  });

  it("propagates non-Not Found metadata lookup errors during publish", async () => {
    const originalFetch = globalThis.fetch;
    const fetchMock = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = String(input);
      const method = String(init?.method || "GET").toUpperCase();

      if (url.endsWith("/git/ref/heads/main") && method === "GET") {
        return new Response(JSON.stringify({ object: { sha: "base-sha-1" } }), { status: 200 });
      }
      if (url.endsWith("/git/refs") && method === "POST") {
        return new Response(JSON.stringify({ ref: "refs/heads/tlk-forge/test" }), { status: 201 });
      }
      if (url.includes("/contents/csv-latest/tlc.csv") && method === "GET") {
        return new Response(JSON.stringify({ message: "Forbidden" }), { status: 403 });
      }
      return new Response(JSON.stringify({ message: `Unhandled ${method} ${url}` }), { status: 500 });
    });

    (globalThis as { fetch: typeof fetch }).fetch = fetchMock as unknown as typeof fetch;
    try {
      await expect(
        publishCsvToGitHub({
          token: "test-token",
          repoFullName: "enonwow/nwn-localization-test",
          baseBranch: "main",
          csvFolder: "csv-latest",
          fileName: "tlc.csv",
          csvBytes: new TextEncoder().encode("A,B\r\n1,2"),
        }),
      ).rejects.toThrow("Forbidden");
    } finally {
      (globalThis as { fetch: typeof fetch }).fetch = originalFetch;
    }
  });
});

