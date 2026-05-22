import type { WorkflowScope } from "../features/workflow/types";

type ScopeSwitcherProps = {
  scope: WorkflowScope;
  onScopeChange: (scope: WorkflowScope) => void;
};

const scopes: Array<{ value: WorkflowScope; label: string }> = [
  { value: "exchange", label: "Exchange" },
  { value: "rebuild", label: "Rebuild" },
];

const ScopeSwitcher = ({ scope, onScopeChange }: ScopeSwitcherProps) => {
  return (
    <div className="scope-switch">
      <p className="scope-switch__label">Workflow Scope</p>
      <div className="scope-switch__tabs" role="tablist" aria-label="Workflow scope">
        {scopes.map((item) => {
          return (
            <button
              key={item.value}
              type="button"
              className={`scope-switch__tab ${scope === item.value ? "scope-switch__tab--active" : ""}`}
              onClick={() => onScopeChange(item.value)}
              aria-selected={scope === item.value}
              role="tab"
            >
              {item.label}
            </button>
          );
        })}
      </div>
    </div>
  );
};

export default ScopeSwitcher;
