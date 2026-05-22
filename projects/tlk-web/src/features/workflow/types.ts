export type WorkflowScope = "exchange" | "rebuild";

export type ImportMode = "tlk" | "xlsx" | "repo";

export interface WorkflowStep {
  title: string;
  sub: string;
}
