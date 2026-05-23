import { describe, expect, it } from "vitest";

import { computeStepUiStates } from "../workflowProgress";

describe("workflowProgress", () => {
  it("locks exchange steps until prerequisites are met", () => {
    const steps = computeStepUiStates("exchange", 0, {
      sourceLoaded: false,
      validated: false,
      exported: false,
      imported: false,
      hasPublishableChanges: false,
    });

    expect(steps).toHaveLength(3);
    expect(steps[0].status).toBe("active");
    expect(steps[1].status).toBe("locked");
    expect(steps[1].canNavigate).toBe(false);
    expect(steps[1].reason).toContain("Load source");
    expect(steps[2].status).toBe("locked");
  });

  it("unlocks publish after source load in exchange scope (validate optional)", () => {
    const steps = computeStepUiStates("exchange", 1, {
      sourceLoaded: true,
      validated: false,
      exported: false,
      imported: false,
      hasPublishableChanges: true,
    });

    expect(steps[0].status).toBe("done");
    expect(steps[1].status).toBe("active");
    expect(steps[2].status).toBe("ready");
    expect(steps[2].canNavigate).toBe(true);
  });

  it("keeps previous and current steps navigable even if flags regress", () => {
    const steps = computeStepUiStates("exchange", 2, {
      sourceLoaded: true,
      validated: false,
      exported: false,
      imported: false,
      hasPublishableChanges: false,
    });

    expect(steps[0].canNavigate).toBe(true);
    expect(steps[1].canNavigate).toBe(true);
    expect(steps[2].canNavigate).toBe(true);
  });

  it("uses rebuild-specific gate for step 2", () => {
    const locked = computeStepUiStates("rebuild", 0, {
      sourceLoaded: false,
      validated: false,
      exported: false,
      imported: false,
      hasPublishableChanges: false,
    });
    expect(locked[1].status).toBe("locked");
    expect(locked[1].reason).toContain("Import approved CSV");

    const ready = computeStepUiStates("rebuild", 0, {
      sourceLoaded: false,
      validated: false,
      exported: false,
      imported: true,
      hasPublishableChanges: false,
    });
    expect(ready[1].status).toBe("ready");
    expect(ready[1].canNavigate).toBe(true);
  });

  it("keeps publish locked when source is loaded but no data changes exist", () => {
    const steps = computeStepUiStates("exchange", 1, {
      sourceLoaded: true,
      validated: false,
      exported: false,
      imported: false,
      hasPublishableChanges: false,
    });

    expect(steps[2].status).toBe("locked");
    expect(steps[2].canNavigate).toBe(false);
    expect(steps[2].reason).toContain("No data changes");
  });
});
