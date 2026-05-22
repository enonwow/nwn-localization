export type WorkflowScope = "exchange" | "rebuild";

export type StepVisualStatus = "done" | "active" | "ready" | "locked";

export interface WorkflowProgressFlags {
  sourceLoaded: boolean;
  validated: boolean;
  exported: boolean;
  imported: boolean;
}

export interface StepUiState {
  status: StepVisualStatus;
  canNavigate: boolean;
  reason?: string;
}

interface StepGate {
  blocked: boolean;
  reason?: string;
}

function buildStepGates(scope: WorkflowScope, flags: WorkflowProgressFlags): StepGate[] {
  if (scope === "exchange") {
    return [
      { blocked: false },
      { blocked: !flags.sourceLoaded, reason: "Load source and run Parse & Validate first." },
      { blocked: !flags.validated, reason: "Validate Edit step first." },
    ];
  }

  return [
    { blocked: false },
    { blocked: !flags.imported, reason: "Import approved CSV first." },
  ];
}

export function computeStepUiStates(
  scope: WorkflowScope,
  activeStep: number,
  flags: WorkflowProgressFlags,
): StepUiState[] {
  const gates = buildStepGates(scope, flags);
  const safeActiveStep = Math.min(Math.max(activeStep, 0), Math.max(0, gates.length - 1));

  return gates.map((gate, index) => {
    const canNavigate = index <= safeActiveStep || !gate.blocked;
    let status: StepVisualStatus;

    if (index === safeActiveStep) {
      status = "active";
    } else if (index < safeActiveStep) {
      status = "done";
    } else if (!gate.blocked) {
      status = "ready";
    } else {
      status = "locked";
    }

    return {
      status,
      canNavigate,
      reason: !canNavigate ? gate.reason : undefined,
    };
  });
}
