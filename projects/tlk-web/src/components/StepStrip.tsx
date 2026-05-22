type StepStripProps = {
  steps: string[];
  activeStep: number;
  stepStates?: Array<{
    status: "done" | "active" | "ready" | "locked";
    canNavigate: boolean;
    reason?: string;
  }>;
  onStepSelect: (index: number) => void;
};

const StepStrip = ({ steps, activeStep, stepStates, onStepSelect }: StepStripProps) => {
  return (
    <nav className="step-strip" aria-label="Workflow steps">
      {steps.map((step, index) => {
        const visualState = stepStates?.[index];
        const status = visualState?.status ?? (index === activeStep ? "active" : "ready");
        const isLocked = visualState ? !visualState.canNavigate : false;

        return (
          <button
            key={step}
            type="button"
            className={`step-strip__step step-strip__step--${status} ${index === activeStep ? "step-strip__step--active" : ""}`}
            onClick={() => onStepSelect(index)}
            disabled={isLocked}
            title={isLocked ? visualState?.reason || "Step is locked." : undefined}
            aria-disabled={isLocked}
          >
            <span className="step-strip__index">{index + 1}</span>
            <span>{step}</span>
            <small className="step-strip__status">{status}</small>
          </button>
        );
      })}
    </nav>
  );
};

export default StepStrip;
