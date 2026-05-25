export interface SetTableCellPlan {
  id: string
  type: 'set_table_cell_value'
  summary: string
  sheetName: string
  tableRange: string
  targetCell: string
  currentValue: unknown
  currentFormula: string | null
  nextValue: unknown
}

export interface SyncTablePlan {
  id: string
  type: 'sync_table_to_table'
  summary: string
  sourceSheetName: string
  sourceTableRange: string
  targetSheetName: string
  targetTableRange: string
  args: Record<string, unknown>
  sourceSignature: string
  targetSignature: string
}

export type OperationPlan = SetTableCellPlan | SyncTablePlan

export class UniverPlanStore {
  private readonly plans = new Map<string, OperationPlan>()

  save<T extends OperationPlan>(plan: T): T {
    this.plans.set(plan.id, plan)
    return plan
  }

  get(planId: string): OperationPlan | null {
    return this.plans.get(planId) ?? null
  }

  delete(planId: string): void {
    this.plans.delete(planId)
  }
}

export function createPlanId(): string {
  if (typeof crypto !== 'undefined' && 'randomUUID' in crypto) {
    return crypto.randomUUID()
  }

  return `plan_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`
}
