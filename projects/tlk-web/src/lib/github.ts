export interface GitHubPublishCsvInput {
  token: string;
  repoFullName: string;
  baseBranch: string;
  csvFolder: string;
  fileName: string;
  csvBytes: Uint8Array;
}

export interface GitHubPublishCsvResult {
  prUrl: string;
  branchName: string;
  filePath: string;
  commitSha: string;
}

interface GitHubApiErrorPayload {
  message?: string;
  errors?: Array<{ message?: string } | string>;
}

interface GitHubRefResponse {
  object: { sha: string };
}

interface GitHubContentResponse {
  sha: string;
}

interface GitHubPutContentResponse {
  commit: { sha: string };
}

interface GitHubPullResponse {
  html_url: string;
}

const GITHUB_API_BASE = "https://api.github.com";

function repoParts(repoFullName: string): { owner: string; repo: string } {
  const normalized = String(repoFullName || "").trim().replace(/^\/+|\/+$/g, "");
  const parts = normalized.split("/").filter(Boolean);
  if (parts.length < 2) {
    throw new Error("Invalid GitHub repo. Use owner/repo.");
  }
  return { owner: parts[0], repo: parts[1] };
}

function normalizePath(value: string): string {
  return String(value || "").trim().replace(/^\/+|\/+$/g, "");
}

function encodeBranchForRef(branch: string): string {
  return branch
    .split("/")
    .filter(Boolean)
    .map((segment) => encodeURIComponent(segment))
    .join("/");
}

function joinPath(left: string, right: string): string {
  const l = normalizePath(left);
  const r = normalizePath(right);
  if (!l) return r;
  if (!r) return l;
  return `${l}/${r}`;
}

function toBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let i = 0; i < bytes.length; i += 1) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

function sanitizeBranchToken(value: string): string {
  const fileStem = String(value || "").replace(/\.[^.]+$/, "");
  const token = fileStem
    .toLowerCase()
    .replace(/[^a-z0-9._/-]+/g, "-")
    .replace(/^-+|-+$/g, "");
  return token || "csv";
}

async function githubRequest<T>(path: string, token: string, init?: RequestInit): Promise<T> {
  const headers = new Headers(init?.headers || {});
  headers.set("Accept", "application/vnd.github+json");
  headers.set("Authorization", `Bearer ${token}`);
  headers.set("X-GitHub-Api-Version", "2022-11-28");
  if (init?.body && !headers.has("Content-Type")) {
    headers.set("Content-Type", "application/json");
  }

  const response = await fetch(`${GITHUB_API_BASE}${path}`, { ...init, headers });
  if (!response.ok) {
    let message = `GitHub API ${response.status}`;
    try {
      const payload = (await response.json()) as GitHubApiErrorPayload;
      const nestedErrors = Array.isArray(payload.errors)
        ? payload.errors
          .map((item) => (typeof item === "string" ? item : item?.message || ""))
          .filter(Boolean)
          .join("; ")
        : "";
      message = [payload.message, nestedErrors].filter(Boolean).join(" | ") || message;
    } catch {
      // keep fallback message
    }
    throw new Error(message);
  }

  if (response.status === 204) {
    return undefined as T;
  }
  return (await response.json()) as T;
}

async function getContentSha(
  owner: string,
  repo: string,
  path: string,
  branch: string,
  token: string,
): Promise<string | null> {
  const encodedPath = path.split("/").map((segment) => encodeURIComponent(segment)).join("/");
  const query = `?ref=${encodeURIComponent(branch)}`;
  try {
    const payload = await githubRequest<GitHubContentResponse>(
      `/repos/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}/contents/${encodedPath}${query}`,
      token,
      { method: "GET" },
    );
    return payload.sha || null;
  } catch (error) {
    const message = error instanceof Error ? error.message : "";
    if (message.includes("Not Found")) {
      return null;
    }
    throw error;
  }
}

async function createBranchFromBase(
  owner: string,
  repo: string,
  baseBranch: string,
  branchName: string,
  token: string,
): Promise<void> {
  const ref = await githubRequest<GitHubRefResponse>(
    `/repos/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}/git/ref/heads/${encodeBranchForRef(baseBranch)}`,
    token,
    { method: "GET" },
  );

  await githubRequest(
    `/repos/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}/git/refs`,
    token,
    {
      method: "POST",
      body: JSON.stringify({
        ref: `refs/heads/${branchName}`,
        sha: ref.object.sha,
      }),
    },
  );
}

export async function publishCsvToGitHub(input: GitHubPublishCsvInput): Promise<GitHubPublishCsvResult> {
  const token = String(input.token || "").trim();
  if (!token) {
    throw new Error("GitHub token is missing.");
  }

  const { owner, repo } = repoParts(input.repoFullName);
  const baseBranch = normalizePath(input.baseBranch) || "main";
  const csvFolder = normalizePath(input.csvFolder) || "csv-latest";
  const fileName = normalizePath(input.fileName);
  if (!fileName) {
    throw new Error("CSV file name is missing.");
  }

  const filePath = joinPath(csvFolder, fileName);
  const now = Date.now().toString(36);
  const branchName = `tlk-forge/${sanitizeBranchToken(fileName)}-${now}`;

  await createBranchFromBase(owner, repo, baseBranch, branchName, token);
  const currentSha = await getContentSha(owner, repo, filePath, branchName, token);
  const encodedPath = filePath.split("/").map((segment) => encodeURIComponent(segment)).join("/");

  const updatePayload = await githubRequest<GitHubPutContentResponse>(
    `/repos/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}/contents/${encodedPath}`,
    token,
    {
      method: "PUT",
      body: JSON.stringify({
        message: `TLK Forge: update ${fileName}`,
        content: toBase64(input.csvBytes),
        branch: branchName,
        ...(currentSha ? { sha: currentSha } : {}),
      }),
    },
  );

  const prPayload = await githubRequest<GitHubPullResponse>(
    `/repos/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}/pulls`,
    token,
    {
      method: "POST",
      body: JSON.stringify({
        title: `TLK Forge: publish ${fileName}`,
        head: branchName,
        base: baseBranch,
        body: `Automated CSV publish from TLK Forge.\n\n- file: ${filePath}`,
      }),
    },
  );

  return {
    prUrl: prPayload.html_url,
    branchName,
    filePath,
    commitSha: updatePayload.commit.sha,
  };
}

