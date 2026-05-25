const WORKBOOK_KEY = 'univer:workbook'
const DB_NAME = 'my-univer-agent'
const DB_VERSION = 1
const STORE_NAME = 'workbook_snapshots'
const ACTIVE_SNAPSHOT_KEY = 'active'

export interface PersistedSheet {
  name: string
  sheetId?: string
  sourceFileId?: string
  sourceFileName?: string
  originalSheetName?: string
  values: (string | number | boolean | null)[][]
  formulas: string[][]
  styles?: Array<Array<Record<string, unknown> | null>>
  columnWidths?: Record<string, number>
  rowHeights?: Record<string, number>
  mergedRanges?: string[]
  hiddenRows?: PersistedIndexRange[]
  hiddenColumns?: PersistedIndexRange[]
  sheetView?: PersistedSheetView
  importStats?: PersistedImportSummary
}

export interface PersistedFileGroup {
  id: string
  name: string
  fileName: string
  createdAt: string
}

export interface PersistedImportSummary {
  sheets?: number
  rows?: number
  columns?: number
  cells?: number
  formulas?: number
  styledCells?: number
  mergedRanges?: number
  hiddenRows?: number
  hiddenColumns?: number
}

export interface PersistedIndexRange {
  start: number
  end: number
}

export interface PersistedSheetView {
  frozenRows?: number
  frozenColumns?: number
  hiddenGridlines?: boolean
  gridlinesColor?: string
  tabColor?: string
}

export interface PersistedWorkbook {
  version: 'grid-v1'
  name: string
  fileName?: string
  activeFileId?: string
  fileGroups?: PersistedFileGroup[]
  importSummary?: PersistedImportSummary
  activeSheetName: string
  sheets: PersistedSheet[]
}

export interface WorkbookSnapshotStats {
  sheets: number
  fileGroups: number
  styledCells: number
  mergedRanges: number
  bytes: number
}

export interface WorkbookStorageResult {
  ok: boolean
  storage: 'indexeddb' | 'localStorage' | 'none'
  stats: WorkbookSnapshotStats
  error?: string
}

interface SaveWorkbookOptions {
  allowStyleLoss?: boolean
}

interface SnapshotRecord {
  id: string
  updatedAt: string
  workbook: PersistedWorkbook
}

function isPersistedWorkbook(value: unknown): value is PersistedWorkbook {
  if (!value || typeof value !== 'object') return false
  const data = value as Partial<PersistedWorkbook>
  return data.version === 'grid-v1' && typeof data.name === 'string' && Array.isArray(data.sheets)
}

export async function saveWorkbook(data: PersistedWorkbook, options: SaveWorkbookOptions = {}): Promise<WorkbookStorageResult> {
  const stats = getWorkbookSnapshotStats(data)
  const record: SnapshotRecord = {
    id: ACTIVE_SNAPSHOT_KEY,
    updatedAt: new Date().toISOString(),
    workbook: data,
  }

  try {
    const database = await openWorkbookDatabase()
    if (!options.allowStyleLoss) {
      const existingRecord = await getSnapshotRecord(database)
      const existingStats = existingRecord?.workbook ? getWorkbookSnapshotStats(existingRecord.workbook) : null
      if (isSuspiciousStyleLoss(existingStats, stats)) {
        database.close()
        return {
          ok: false,
          storage: 'indexeddb',
          stats,
          error: `Refusing to overwrite styled snapshot (${existingStats?.styledCells ?? 0} styled cells) with a style-less snapshot.`,
        }
      }
    }
    await putSnapshotRecord(database, record)
    database.close()
    return { ok: true, storage: 'indexeddb', stats }
  } catch (error) {
    try {
      localStorage.setItem(WORKBOOK_KEY, JSON.stringify(data))
      return { ok: true, storage: 'localStorage', stats, error: getErrorMessage(error) }
    } catch (fallbackError) {
      return {
        ok: false,
        storage: 'none',
        stats,
        error: `${getErrorMessage(error)}; fallback failed: ${getErrorMessage(fallbackError)}`,
      }
    }
  }
}

function isSuspiciousStyleLoss(previous: WorkbookSnapshotStats | null, next: WorkbookSnapshotStats): boolean {
  return !!previous &&
    previous.styledCells > 0 &&
    (next.styledCells === 0 || next.styledCells < Math.max(8, Math.floor(previous.styledCells * 0.25))) &&
    next.sheets >= previous.sheets
}

export async function loadWorkbook(): Promise<PersistedWorkbook | null> {
  try {
    const database = await openWorkbookDatabase()
    const record = await getSnapshotRecord(database)
    database.close()
    if (isPersistedWorkbook(record?.workbook)) return record.workbook
  } catch {
    // Fall back to legacy localStorage below.
  }

  const legacyWorkbook = loadLegacyWorkbook()
  if (legacyWorkbook) {
    void saveWorkbook(legacyWorkbook)
  }
  return legacyWorkbook
}

export async function clearWorkbook(): Promise<void> {
  try {
    const database = await openWorkbookDatabase()
    await deleteSnapshotRecord(database)
    database.close()
  } catch {
    // localStorage cleanup below is still useful if IndexedDB is unavailable.
  }
  localStorage.removeItem(WORKBOOK_KEY)
}

export function getWorkbookSnapshotStats(data: PersistedWorkbook): WorkbookSnapshotStats {
  return {
    sheets: Array.isArray(data.sheets) ? data.sheets.length : 0,
    fileGroups: Array.isArray(data.fileGroups) ? data.fileGroups.length : 0,
    styledCells: countStyledCells(data),
    mergedRanges: countMergedRanges(data),
    bytes: estimateJsonBytes(data),
  }
}

function loadLegacyWorkbook(): PersistedWorkbook | null {
  try {
    const raw = localStorage.getItem(WORKBOOK_KEY)
    if (!raw) return null
    const parsed = JSON.parse(raw)
    if (!isPersistedWorkbook(parsed)) {
      localStorage.removeItem(WORKBOOK_KEY)
      return null
    }
    return parsed
  } catch {
    return null
  }
}

function openWorkbookDatabase(): Promise<IDBDatabase> {
  if (!window.indexedDB) {
    return Promise.reject(new Error('IndexedDB is not available'))
  }

  return new Promise((resolve, reject) => {
    const request = window.indexedDB.open(DB_NAME, DB_VERSION)

    request.onupgradeneeded = () => {
      const database = request.result
      if (!database.objectStoreNames.contains(STORE_NAME)) {
        database.createObjectStore(STORE_NAME, { keyPath: 'id' })
      }
    }

    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error ?? new Error('Failed to open IndexedDB'))
    request.onblocked = () => reject(new Error('IndexedDB upgrade is blocked'))
  })
}

function putSnapshotRecord(database: IDBDatabase, record: SnapshotRecord): Promise<void> {
  return runSnapshotTransaction(database, 'readwrite', (store) => store.put(record))
}

function getSnapshotRecord(database: IDBDatabase): Promise<SnapshotRecord | null> {
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(STORE_NAME, 'readonly')
    const store = transaction.objectStore(STORE_NAME)
    const request = store.get(ACTIVE_SNAPSHOT_KEY)
    let result: SnapshotRecord | null = null

    request.onsuccess = () => {
      result = request.result as SnapshotRecord | undefined ?? null
    }
    request.onerror = () => reject(request.error ?? new Error('Failed to read workbook snapshot'))
    transaction.oncomplete = () => resolve(result)
    transaction.onerror = () => reject(transaction.error ?? new Error('Workbook snapshot read transaction failed'))
    transaction.onabort = () => reject(transaction.error ?? new Error('Workbook snapshot read transaction aborted'))
  })
}

function deleteSnapshotRecord(database: IDBDatabase): Promise<void> {
  return runSnapshotTransaction(database, 'readwrite', (store) => store.delete(ACTIVE_SNAPSHOT_KEY))
}

function runSnapshotTransaction(
  database: IDBDatabase,
  mode: IDBTransactionMode,
  run: (store: IDBObjectStore) => IDBRequest,
): Promise<void> {
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(STORE_NAME, mode)
    const store = transaction.objectStore(STORE_NAME)
    const request = run(store)

    request.onerror = () => reject(request.error ?? new Error('Workbook snapshot request failed'))
    transaction.oncomplete = () => resolve()
    transaction.onerror = () => reject(transaction.error ?? new Error('Workbook snapshot transaction failed'))
    transaction.onabort = () => reject(transaction.error ?? new Error('Workbook snapshot transaction aborted'))
  })
}

function countStyledCells(data: PersistedWorkbook): number {
  return (data.sheets ?? []).reduce((total, sheet) => {
    const styles = Array.isArray(sheet.styles) ? sheet.styles : []
    return total + styles.reduce((rowTotal, row) =>
      rowTotal + (Array.isArray(row) ? row.filter(Boolean).length : 0), 0)
  }, 0)
}

function countMergedRanges(data: PersistedWorkbook): number {
  return (data.sheets ?? []).reduce((total, sheet) =>
    total + (Array.isArray(sheet.mergedRanges) ? sheet.mergedRanges.length : 0), 0)
}

function estimateJsonBytes(data: PersistedWorkbook): number {
  try {
    return new Blob([JSON.stringify(data)]).size
  } catch {
    return JSON.stringify(data).length
  }
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error)
}
