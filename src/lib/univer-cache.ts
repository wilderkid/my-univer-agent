import type { WorkbookLike, WorkbookSchemaSnapshot } from './univer-schema'
import { extractWorkbookSchema } from './univer-schema'

interface CacheEntry {
  workbook: WorkbookLike
  snapshot: WorkbookSchemaSnapshot
}

export class UniverSchemaCache {
  private entry: CacheEntry | null = null

  getSnapshot(workbook: WorkbookLike): WorkbookSchemaSnapshot {
    if (!this.entry || this.entry.workbook !== workbook) {
      this.entry = {
        workbook,
        snapshot: extractWorkbookSchema(workbook),
      }
    }

    return this.entry.snapshot
  }

  invalidate(workbook?: WorkbookLike): void {
    if (!this.entry) return
    if (!workbook || this.entry.workbook === workbook) {
      this.entry = null
    }
  }

  refresh(workbook: WorkbookLike): WorkbookSchemaSnapshot {
    this.entry = {
      workbook,
      snapshot: extractWorkbookSchema(workbook),
    }
    return this.entry.snapshot
  }
}
