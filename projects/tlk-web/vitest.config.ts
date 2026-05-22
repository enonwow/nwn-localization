import { defineConfig, mergeConfig } from "vitest/config";
import viteConfig from "./vite.config";

export default mergeConfig(
  viteConfig,
  defineConfig({
    test: {
      include: ["src/lib/__tests__/**/*.test.ts"],
      coverage: {
        provider: "v8",
        include: ["src/lib/**/*.ts"],
        exclude: ["src/lib/**/__tests__/**"],
        thresholds: {
          statements: 85,
          lines: 85,
          functions: 85,
          branches: 80,
        },
      },
    },
  }),
);
