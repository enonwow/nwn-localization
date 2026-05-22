/** @type {import('@stryker-mutator/api/core').PartialStrykerOptions} */
export default {
  packageManager: "npm",
  testRunner: "vitest",
  reporters: ["progress", "clear-text", "html"],
  mutate: [
    "src/lib/**/*.ts",
    "!src/lib/**/__tests__/**",
    "!src/lib/types.ts",
  ],
  checkers: ["typescript"],
  tsconfigFile: "tsconfig.json",
  coverageAnalysis: "off",
  concurrency: 4,
  timeoutMS: 120000,
  timeoutFactor: 1.5,
  thresholds: {
    high: 80,
    low: 70,
    break: 65,
  },
  vitest: {
    configFile: "vitest.config.ts",
  },
};
