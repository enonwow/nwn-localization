import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

const githubPagesBase =
  process.env.GITHUB_PAGES_BASE?.trim() ||
  (process.env.GITHUB_ACTIONS === "true" && process.env.GITHUB_REPOSITORY
    ? `/${process.env.GITHUB_REPOSITORY.split("/").pop()}/`
    : "/");

export default defineConfig({
  plugins: [react()],
  base: githubPagesBase,
});
