const WORKBOOK_KEY = 'univer:workbook'
const AI_CONFIG_KEY = 'univer:ai-config'

// ── Workbook ─────────────────────────────────────────────────────────────────

export interface PersistedSheet {
  name: string
  values: (string | number | boolean | null)[][]
  formulas: string[][]
}

export interface PersistedWorkbook {
  version: 'grid-v1'
  name: string
  activeSheetName: string
  sheets: PersistedSheet[]
}

function isPersistedWorkbook(value: unknown): value is PersistedWorkbook {
  if (!value || typeof value !== 'object') return false
  const data = value as Partial<PersistedWorkbook>
  return data.version === 'grid-v1' && typeof data.name === 'string' && Array.isArray(data.sheets)
}

export function saveWorkbook(data: PersistedWorkbook): void {
  try {
    localStorage.setItem(WORKBOOK_KEY, JSON.stringify(data))
  } catch {
    // localStorage 满或不可用时静默忽略
  }
}

export function loadWorkbook(): PersistedWorkbook | null {
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

export function clearWorkbook(): void {
  localStorage.removeItem(WORKBOOK_KEY)
}

// ── AI config ─────────────────────────────────────────────────────────────────

export interface SavedAIConfig {
  provider: string
  apiKey: string
  model: string
  baseURL: string
}

export function saveAIConfig(config: SavedAIConfig): void {
  try {
    localStorage.setItem(AI_CONFIG_KEY, JSON.stringify(config))
  } catch { /* ignore */ }
}

export function loadAIConfig(): SavedAIConfig | null {
  try {
    const raw = localStorage.getItem(AI_CONFIG_KEY)
    return raw ? (JSON.parse(raw) as SavedAIConfig) : null
  } catch {
    return null
  }
}
