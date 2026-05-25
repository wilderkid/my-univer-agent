import { useCallback, useEffect, useRef, useState, type ChangeEvent, type CSSProperties, type KeyboardEvent } from 'react'
import type { LLMProvider } from '../lib/agent'
import { destroyAgent, getAgent, initAgent } from '../lib/agent'
import {
  clearBackendAIConfig,
  deleteBackendAIModel,
  discoverBackendAIModels,
  loadBackendAISettings,
  saveBackendAIModels,
  selectActiveBackendAIModel,
  type BackendAIModel,
  type BackendAISettings,
} from '../lib/ai-config'
import type { AgentApplyPlanResult, AgentExecutionResult, AgentPlanPreview, AgentSnapshotDiff, AgentSnapshotSummary, UniverAgent } from '../lib/univer-agent'
import type { AgentToolExecution } from '../lib/univer-agent'
import { getUniverAPI } from '../lib/univer-ref'
import type { RangeLike, SheetLike, WorkbookLike } from '../lib/univer-schema'
import type { AuthUser } from '../lib/auth'

interface VerificationSummary {
  passed: boolean | null
  text: string | null
}

type UnknownRecord = Record<string, unknown>
type ExecutionStatus = 'running' | 'success' | 'warning' | 'error' | 'undone'
type UndoState = 'idle' | 'undoing' | 'undone' | 'failed'
type DiffFilter = 'all' | 'value' | 'formula' | 'format'

interface ExecutionRecord {
  id: number
  instruction: string
  status: ExecutionStatus
  message: string
  startedAt: number
  finishedAt?: number
  tools: AgentToolExecution[]
  plans: AgentPlanPreview[]
  applyResult: AgentApplyPlanResult | null
  verification: VerificationSummary
  snapshot: AgentSnapshotSummary | null
  afterSnapshot: AgentSnapshotSummary | null
  diff: AgentSnapshotDiff | null
  undoSteps: number
  undoState: UndoState
  error?: string
}

interface SheetPickerState {
  tokenStart: number
  tokenEnd: number
  query: string
}

interface RangeBindingState {
  start: number
  end: number
  value: string
}

const EMPTY_SETTINGS: BackendAISettings = {
  activeModelId: '',
  models: [],
}

const MAX_EXECUTION_RECORDS = 50
const DEFAULT_VISIBLE_RECORD_LIMIT = 8
const VISIBLE_RECORD_INCREMENT = 8

const VALID_PROVIDERS: LLMProvider[] = ['qwen', 'openai', 'custom']

const MUTATING_TOOL_NAMES = new Set([
  'set_table_cell_value',
  'sync_table_to_table',
  'apply_plan',
  'append_table_records',
  'set_cell_value',
  'set_range_values',
  'format_range',
  'set_freeze',
  'hide_rows',
  'show_rows',
  'hide_columns',
  'show_columns',
  'set_gridlines',
  'set_sheet_tab_color',
  'set_column_width',
  'set_row_height',
  'auto_resize_columns',
  'auto_resize_rows',
  'clear_range',
  'insert_sheet',
  'rename_sheet',
  'insert_rows',
  'delete_rows',
  'insert_columns',
  'delete_columns',
  'merge_cells',
])

const PREVIEW_TOOL_NAMES = new Set([
  'preview_set_table_cell_value',
  'preview_sync_table_to_table',
])

const PROVIDER_LABELS: Record<LLMProvider, string> = {
  qwen: '通义千问',
  openai: 'OpenAI',
  custom: '自定义',
}

const PROVIDER_BASE_URL_PLACEHOLDER: Record<LLMProvider, string> = {
  qwen: '默认 https://dashscope.aliyuncs.com/compatible-mode/v1',
  openai: '默认 https://api.openai.com/v1',
  custom: '例如 http://127.0.0.1:11434/v1',
}

function isRecord(value: unknown): value is UnknownRecord {
  return !!value && typeof value === 'object' && !Array.isArray(value)
}

function asRecord(value: unknown): UnknownRecord | null {
  return isRecord(value) ? value : null
}

function asArray(value: unknown): unknown[] {
  return Array.isArray(value) ? value : []
}

function asString(value: unknown): string | null {
  return typeof value === 'string' && value.trim() ? value.trim() : null
}

function asNumber(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null
}

function formatValue(value: unknown): string {
  if (value === null || value === undefined || value === '') return '空'
  if (typeof value === 'string') return value
  if (typeof value === 'number' || typeof value === 'boolean') return String(value)
  try {
    return JSON.stringify(value)
  } catch {
    return String(value)
  }
}

function formatDerivedReason(reason: string): string {
  return reason
    .replace(/^Derived from /, '来自 ')
    .replace(' / ', ' ÷ ')
    .replace(' * ', ' × ')
}

function columnNumberToLetter(columnNumber: number): string {
  let result = ''
  let current = columnNumber

  while (current > 0) {
    const remainder = (current - 1) % 26
    result = String.fromCharCode(65 + remainder) + result
    current = Math.floor((current - 1) / 26)
  }

  return result || 'A'
}

function getPlanPreview(plan: AgentPlanPreview): UnknownRecord | null {
  return asRecord(plan.result.preview)
}

function getResultVerification(result: unknown): UnknownRecord | null {
  const record = asRecord(result)
  return record ? asRecord(record.verification) : null
}

function summarizeMismatch(mismatch: unknown): string | null {
  const data = asRecord(mismatch)
  if (!data) return null

  const rowNumber = asNumber(data.row_number)
  const columnNumber = asNumber(data.column_number)
  const expectedValue = data.expected_value ?? null
  const actualValue = data.actual_value ?? null
  const actualFormula = asString(data.actual_formula)
  const cellRef = rowNumber && columnNumber ? `${columnNumberToLetter(columnNumber)}${rowNumber}` : '未知单元格'

  return `${cellRef}: 期望 ${formatValue(expectedValue)}，实际 ${formatValue(actualFormula ?? actualValue)}`
}

function getVerificationSummary(result: unknown): VerificationSummary {
  const verification = getResultVerification(result)
  if (!verification) return { passed: null, text: null }

  if ('updated_row_checks' in verification || 'appended_row_check' in verification) {
    const updatedChecks = asArray(verification.updated_row_checks).map(asRecord).filter((item): item is UnknownRecord => !!item)
    const failedUpdatedChecks = updatedChecks.filter((entry) => entry.verified === false)
    const appendedCheck = asRecord(verification.appended_row_check)
    const appendedMismatches = appendedCheck ? asArray(appendedCheck.mismatches).map(summarizeMismatch).filter(Boolean) : []
    const updateMismatches = failedUpdatedChecks.flatMap((entry) =>
      asArray(entry.mismatches).map(summarizeMismatch).filter(Boolean),
    )
    const passed = typeof verification.verified === 'boolean' ? verification.verified : null

    return {
      passed,
      text: [
        passed === false ? '写后核对未通过。' : '写后核对通过。',
        `更新核对批次：${updatedChecks.length}，失败：${failedUpdatedChecks.length}`,
        appendedCheck ? `追加核对：${appendedCheck.verified ? '通过' : '失败'}` : '追加核对：无',
        ...updateMismatches.slice(0, 8),
        ...appendedMismatches.slice(0, 8),
      ].join('\n'),
    }
  }

  if ('mismatches' in verification || 'mismatch_count' in verification) {
    const passed = typeof verification.verified === 'boolean' ? verification.verified : null
    const mismatchLines = asArray(verification.mismatches).map(summarizeMismatch).filter(Boolean)

    return {
      passed,
      text: [
        passed === false ? '写后核对未通过。' : '写后核对通过。',
        typeof verification.mismatch_count === 'number' ? `不一致单元格：${verification.mismatch_count}` : null,
        asString(verification.error),
        ...mismatchLines.slice(0, 8),
      ].filter(Boolean).join('\n'),
    }
  }

  const passed = typeof verification.verified === 'boolean' ? verification.verified : null
  return {
    passed,
    text: [
      passed === false ? '写后核对未通过。' : '写后核对通过。',
      `期望：${formatValue(verification.expected_value ?? null)}`,
      `实际：${formatValue(verification.actual_formula ?? verification.actual_value ?? null)}`,
      asString(verification.error),
    ].filter(Boolean).join('\n'),
  }
}

function summarizePlanPreview(plan: AgentPlanPreview): string {
  const preview = getPlanPreview(plan)
  if (!preview) return plan.summary

  const appendedRows = asNumber(preview.appended_rows) ?? 0
  const updatedRows = asArray(preview.updated_rows).length
  const skippedRows = asArray(preview.skipped_rows).length
  const unmatchedRows = asArray(preview.unmatched_rows).length
  const ambiguousRows = asArray(preview.ambiguous_rows).length

  if (plan.type === 'sync_table_to_table') {
    return `${plan.summary}。预计新增 ${appendedRows} 行，更新 ${updatedRows} 行，跳过 ${skippedRows} 行，未匹配 ${unmatchedRows} 行，歧义 ${ambiguousRows} 行。`
  }

  const location = asRecord(plan.result.location)
  if (location?.target_cell) {
    return `${plan.summary}。目标单元格：${String(location.target_cell)}。`
  }

  return plan.summary
}

function getPreviewStats(preview: UnknownRecord | null): Array<{ label: string; value: string; tone?: 'warn' | 'bad' }> {
  if (!preview) return []

  return [
    { label: '更新', value: String(asArray(preview.updated_rows).length) },
    { label: '新增', value: String(asNumber(preview.appended_rows) ?? 0) },
    { label: '跳过', value: String(asArray(preview.skipped_rows).length), tone: asArray(preview.skipped_rows).length > 0 ? 'warn' : undefined },
    { label: '未匹配', value: String(asArray(preview.unmatched_rows).length), tone: asArray(preview.unmatched_rows).length > 0 ? 'bad' : undefined },
    { label: '歧义', value: String(asArray(preview.ambiguous_rows).length), tone: asArray(preview.ambiguous_rows).length > 0 ? 'bad' : undefined },
  ]
}

function getMatchRows(preview: UnknownRecord | null): string[] {
  if (!preview) return []

  return asArray(preview.match_headers)
    .map(asRecord)
    .filter((item): item is UnknownRecord => !!item)
    .map((item) => {
      const source = asString(item.source_header) ?? '源字段'
      const target = asString(item.target_header) ?? '目标字段'
      const score = asNumber(item.score)
      return score === null ? `${source} -> ${target}` : `${source} -> ${target}，置信度 ${score}`
    })
}

function getMatchAnalysisRows(preview: UnknownRecord | null): string[] {
  if (!preview) return []

  return asArray(preview.match_header_analysis)
    .map(asRecord)
    .filter((item): item is UnknownRecord => !!item)
    .slice(0, 8)
    .map((item) => {
      const source = asString(item.sourceHeader) ?? '源字段'
      const target = asString(item.targetHeader) ?? '目标字段'
      const sharedValueCount = asNumber(item.sharedValueCount) ?? 0
      const sourceDistinctCount = asNumber(item.sourceDistinctCount) ?? 0
      const targetDistinctCount = asNumber(item.targetDistinctCount) ?? 0
      const samples = asArray(item.sharedValueSamples).map(formatValue).filter(Boolean).slice(0, 5)
      const sampleText = samples.length > 0 ? `；样本：${samples.join(', ')}` : ''

      return `${source} -> ${target}：共享值 ${sharedValueCount}，源唯一值 ${sourceDistinctCount}，目标唯一值 ${targetDistinctCount}${sampleText}`
    })
}

function getDerivedRows(preview: UnknownRecord | null): string[] {
  if (!preview) return []

  return asArray(preview.derived_cells)
    .map(asRecord)
    .filter((item): item is UnknownRecord => !!item)
    .flatMap((item) => {
      const sourceRow = asNumber(item.source_row_number)
      const targetRow = asNumber(item.target_row_number)
      return asArray(item.cells)
        .map(asRecord)
        .filter((cell): cell is UnknownRecord => !!cell)
        .map((cell) => {
          const targetHeader = asString(cell.targetHeader) ?? '目标字段'
          const value = formatValue(cell.value)
          const reason = formatDerivedReason(asString(cell.reason) ?? '派生计算')
          const rowText = targetRow ? `目标第 ${targetRow} 行` : sourceRow ? `源第 ${sourceRow} 行` : '待写入行'
          return `${rowText}：${targetHeader} = ${value}（${reason}）`
        })
    })
}

function getWriteRows(preview: UnknownRecord | null): string[] {
  if (!preview) return []

  const updatedRows = asArray(preview.updated_rows)
    .map(asRecord)
    .filter((item): item is UnknownRecord => !!item)
    .slice(0, 8)
    .map((item) => {
      const rowNumber = asNumber(item.row_number)
      const matchedOn = asArray(item.matched_on).map(formatValue).join(', ')
      const matchScore = asNumber(item.match_score)
      const scoreText = matchScore === null ? '' : `，匹配分 ${matchScore}`
      return `更新目标第 ${rowNumber ?? '?'} 行${scoreText}${matchedOn ? `，依据 ${matchedOn}` : ''}`
    })

  const appendedRange = asString(preview.appended_write_range)
  if (appendedRange) updatedRows.push(`追加写入区域：${appendedRange}`)

  return updatedRows
}

function getIssueRows(preview: UnknownRecord | null): string[] {
  if (!preview) return []

  const skippedRows = asArray(preview.skipped_rows)
    .map(asRecord)
    .filter((item): item is UnknownRecord => !!item)
    .slice(0, 5)
    .map((item) => `跳过源第 ${asNumber(item.source_row_number) ?? '?'} 行：${asString(item.reason) ?? '未知原因'}`)

  const unmatchedRows = asArray(preview.unmatched_rows)
    .map(asRecord)
    .filter((item): item is UnknownRecord => !!item)
    .slice(0, 5)
    .map((item) => `未匹配源第 ${asNumber(item.source_row_number) ?? '?'} 行：${asString(item.reason) ?? '未找到对应目标行'}`)

  const ambiguousRows = asArray(preview.ambiguous_rows)
    .map(asRecord)
    .filter((item): item is UnknownRecord => !!item)
    .slice(0, 5)
    .map((item) => {
      const candidates = asArray(item.candidate_row_numbers).map(formatValue).join(', ')
      return `歧义源第 ${asNumber(item.source_row_number) ?? '?'} 行：候选目标行 ${candidates || '未知'}`
    })

  return [...skippedRows, ...unmatchedRows, ...ambiguousRows]
}

function isMutatingTool(toolName: string): boolean {
  return MUTATING_TOOL_NAMES.has(toolName)
}

function getUndoStepCount(result: AgentExecutionResult, applyResult: AgentApplyPlanResult | null): number {
  const directMutationCount = result.tools.filter((tool) => isMutatingTool(tool.toolName) && tool.toolName !== 'apply_plan').length
  return directMutationCount + (applyResult ? 1 : 0)
}

function getToolBadge(toolName: string): string {
  if (
    toolName.includes('format') ||
    toolName.includes('resize') ||
    toolName.includes('width') ||
    toolName.includes('height') ||
    toolName.includes('freeze') ||
    toolName.includes('gridlines') ||
    toolName.includes('tab_color') ||
    toolName.includes('hide_') ||
    toolName.includes('show_')
  ) return '格式'
  if (toolName.includes('sync') || toolName.includes('append') || toolName.includes('set_')) return '写入'
  if (toolName.includes('merge') || toolName.includes('insert') || toolName.includes('delete') || toolName.includes('sheet')) return '结构'
  if (toolName.includes('preview') || toolName.includes('locate') || toolName.includes('find') || toolName.includes('describe') || toolName.includes('list')) return '分析'
  return '读取'
}

function getToolTitle(toolName: string): string {
  const titles: Record<string, string> = {
    get_spreadsheet_info: '读取工作簿结构',
    list_tables: '识别表格',
    describe_table: '读取表格明细',
    find_rows_by_conditions: '按条件查找行',
    locate_target_cell: '定位目标单元格',
    preview_set_table_cell_value: '预览单元格写入',
    set_table_cell_value: '写入单元格',
    preview_sync_table_to_table: '预览跨表同步',
    sync_table_to_table: '跨表同步',
    apply_plan: '应用执行计划',
    append_table_records: '追加表格记录',
    set_cell_value: '写入指定单元格',
    set_range_values: '写入指定区域',
    get_range_values: '读取指定区域',
    format_range: '设置区域格式',
    set_freeze: '冻结窗格',
    hide_rows: '隐藏行',
    show_rows: '显示行',
    hide_columns: '隐藏列',
    show_columns: '显示列',
    set_gridlines: '设置网格线',
    set_sheet_tab_color: '设置工作表标签色',
    set_column_width: '设置列宽',
    set_row_height: '设置行高',
    auto_resize_columns: '自动调整列宽',
    auto_resize_rows: '自动调整行高',
    clear_range: '清除区域',
    insert_sheet: '新建工作表',
    rename_sheet: '重命名工作表',
    insert_rows: '插入行',
    delete_rows: '删除行',
    insert_columns: '插入列',
    delete_columns: '删除列',
    merge_cells: '合并单元格',
  }

  return titles[toolName] ?? toolName
}

function getToolScope(tool: AgentToolExecution): string {
  const args = tool.args
  const result = asRecord(tool.result)
  const resultRange = asString(result?.range) ?? asString(result?.write_range) ?? asString(result?.appended_write_range)
  const location = asRecord(result?.location)
  const targetCell = asString(location?.target_cell)
  const sheetName = asString(args.sheet_name) ?? asString(args.target_sheet_name) ?? asString(args.source_sheet_name)
  const range = asString(args.range) ?? resultRange ?? targetCell
  const columnIndex = asNumber(args.column_index)
  const rowIndex = asNumber(args.row_index)

  if (range && sheetName) return `${sheetName}!${range}`
  if (range) return range
  if (columnIndex !== null) return `${columnNumberToLetter(columnIndex + 1)} 列`
  if (rowIndex !== null) return `${rowIndex + 1} 行`
  if (sheetName) return sheetName
  return '当前工作区'
}

function getToolResultState(tool: AgentToolExecution): 'ok' | 'warn' | 'error' | 'info' {
  const result = asRecord(tool.result)
  if (result?.error) return 'error'
  if (result?.warning) return 'warn'
  if (PREVIEW_TOOL_NAMES.has(tool.toolName)) return 'info'
  if (isMutatingTool(tool.toolName)) return 'ok'
  return 'info'
}

function getExecutionMetrics(record: ExecutionRecord): Array<{ label: string; value: string; tone?: 'bad' | 'warn' }> {
  const toolCount = record.tools.length
  const mutationCount = record.tools.filter((tool) => isMutatingTool(tool.toolName)).length + (record.applyResult ? 1 : 0)
  const preview = record.plans.at(-1) ? getPlanPreview(record.plans.at(-1) as AgentPlanPreview) : null
  const updatedRows = preview ? asArray(preview.updated_rows).length : 0
  const appendedRows = preview ? asNumber(preview.appended_rows) ?? 0 : 0
  const issueRows = preview ? getIssueRows(preview).length : 0
  const changedCells = record.diff?.success ? record.diff.changedCellCount : 0

  return [
    { label: '步骤', value: String(toolCount) },
    { label: '修改', value: String(mutationCount) },
    { label: '差异', value: String(changedCells), tone: changedCells > 0 ? undefined : 'warn' },
    { label: '更新行', value: String(updatedRows) },
    { label: '新增行', value: String(appendedRows) },
    { label: '问题', value: String(issueRows), tone: issueRows > 0 ? 'warn' : undefined },
    { label: '快照', value: record.snapshot ? String(record.snapshot.sheetCount) : '0' },
  ]
}

function formatRecordTime(timestamp: number): string {
  return new Date(timestamp).toLocaleTimeString('zh-CN', { hour12: false, hour: '2-digit', minute: '2-digit', second: '2-digit' })
}

function getCurrentTimestamp(): number {
  return new Date().getTime()
}

function getDiffKindLabel(kind: string): string {
  if (kind === 'formula') return '公式'
  if (kind === 'value_formula') return '值+公式'
  return '值'
}

function formatDiffCellValue(value: unknown, formula: string | null): string {
  if (formula) return formula
  return formatValue(value)
}

function getDisplayableDiffChanges(diff: AgentSnapshotDiff) {
  return diff.changes.filter((change) => change.kind !== 'format')
}

function safeRead<T>(reader: () => T): T | null {
  try {
    return reader()
  } catch {
    return null
  }
}

function getWorkbookFacade(): WorkbookLike | null {
  const api = getUniverAPI() as unknown as { getActiveWorkbook?: () => WorkbookLike | null } | null
  return api?.getActiveWorkbook?.() ?? null
}

function getWorkbookSheetNames(): string[] {
  const workbook = getWorkbookFacade()
  const sheets = workbook ? safeRead(() => workbook.getSheets()) : null
  return sheets?.map((sheet) => sheet.getSheetName()).filter(Boolean) ?? []
}

function normalizeRangeA1Notation(notation: string): string {
  const withoutSheet = notation.includes('!') ? notation.slice(notation.lastIndexOf('!') + 1) : notation
  return withoutSheet.replace(/\$/g, '')
}

function getSelectionRange(sheet: SheetLike, workbook: WorkbookLike): RangeLike | null {
  const selectedRanges = safeRead(() => sheet.getSelection()?.getActiveRangeList() ?? [])
  if (selectedRanges && selectedRanges.length > 0) return selectedRanges[selectedRanges.length - 1] ?? null

  const activeRange = safeRead(() => workbook.getActiveRange())
  if (activeRange) return activeRange

  const activeCell = safeRead(() => workbook.getActiveCell())
  if (!activeCell) return null

  return {
    getA1Notation: () => activeCell.getA1Notation(),
  } as RangeLike
}

function getActiveSelectionReference(): string | null {
  const workbook = getWorkbookFacade()
  const sheet = workbook ? safeRead(() => workbook.getActiveSheet()) : null
  if (!workbook || !sheet) return null

  const range = getSelectionRange(sheet, workbook)
  const notation = range ? safeRead(() => range.getA1Notation(false)) : null
  if (!notation) return null

  return `${sheet.getSheetName()}:${normalizeRangeA1Notation(notation)}`
}

function isStandaloneTrigger(value: string, triggerIndex: number): boolean {
  const previous = value[triggerIndex - 1]
  return !previous || !/[A-Za-z0-9_]/.test(previous)
}

interface AgentPanelProps {
  currentUser: AuthUser
}

export function AgentPanel({ currentUser }: AgentPanelProps) {
  const canConfigureAI = currentUser.role === 'admin'
  const [assistantOpen, setAssistantOpen] = useState(false)
  const [configOpen, setConfigOpen] = useState(false)
  const [loadingConfig, setLoadingConfig] = useState(true)
  const [savingConfig, setSavingConfig] = useState(false)
  const [discoveringModels, setDiscoveringModels] = useState(false)
  const [switchingModel, setSwitchingModel] = useState(false)
  const [applyingPlan, setApplyingPlan] = useState(false)
  const [provider, setProvider] = useState<LLMProvider>('qwen')
  const [apiKey, setApiKey] = useState('')
  const [baseURL, setBaseURL] = useState('')
  const [manualModel, setManualModel] = useState('')
  const [settings, setSettings] = useState<BackendAISettings>(EMPTY_SETTINGS)
  const [discoveredModels, setDiscoveredModels] = useState<string[]>([])
  const [selectedModels, setSelectedModels] = useState<string[]>([])
  const [configured, setConfigured] = useState(false)
  const [cmd, setCmd] = useState('')
  const [sheetNames, setSheetNames] = useState<string[]>([])
  const [sheetPicker, setSheetPicker] = useState<SheetPickerState | null>(null)
  const [rangeBinding, setRangeBinding] = useState<RangeBindingState | null>(null)
  const [running, setRunning] = useState(false)
  const [status, setStatus] = useState('正在加载 AI 配置...')
  const [executionRecords, setExecutionRecords] = useState<ExecutionRecord[]>([])
  const [visibleRecordLimit, setVisibleRecordLimit] = useState(DEFAULT_VISIBLE_RECORD_LIMIT)
  const [conversationStartId, setConversationStartId] = useState(1)
  const [showArchivedRecords, setShowArchivedRecords] = useState(false)
  const [activeDiffRecord, setActiveDiffRecord] = useState<ExecutionRecord | null>(null)

  const agentRef = useRef<UniverAgent | null>(null)
  const textareaRef = useRef<HTMLTextAreaElement>(null)
  const cmdRef = useRef('')
  const rangeBindingRef = useRef<RangeBindingState | null>(null)
  const logEndRef = useRef<HTMLDivElement>(null)
  const executionIdRef = useRef(1)

  const getActiveModel = useCallback((nextSettings: BackendAISettings): BackendAIModel | null => {
    if (nextSettings.models.length === 0) return null
    return nextSettings.models.find((entry) => entry.id === nextSettings.activeModelId) ?? nextSettings.models[0] ?? null
  }, [])

  const syncFormFromModel = useCallback((model: BackendAIModel | null) => {
    if (!model) return
    const nextProvider = VALID_PROVIDERS.includes(model.provider as LLMProvider)
      ? (model.provider as LLMProvider)
      : 'custom'

    setProvider(nextProvider)
    setApiKey(model.apiKey)
    setBaseURL(model.baseURL)
  }, [])

  const initializeAgent = useCallback((model: BackendAIModel) => {
    const resolvedProvider = VALID_PROVIDERS.includes(model.provider as LLMProvider)
      ? (model.provider as LLMProvider)
      : 'custom'

    agentRef.current?.cancel()
    agentRef.current = initAgent({
      apiKey: model.apiKey.trim(),
      provider: resolvedProvider,
      model: model.model,
      baseURL: model.baseURL,
    })
  }, [])

  const resetAgent = useCallback(() => {
    agentRef.current?.cancel()
    agentRef.current = null
    destroyAgent()
  }, [])

  const updateExecutionRecord = useCallback((id: number, patch: Partial<ExecutionRecord>) => {
    setExecutionRecords((prev) => prev.map((record) => (record.id === id ? { ...record, ...patch } : record)))
  }, [])

  const startNewConversation = useCallback(() => {
    setConversationStartId(executionIdRef.current)
    setVisibleRecordLimit(DEFAULT_VISIBLE_RECORD_LIMIT)
    setShowArchivedRecords(false)
    setStatus('已开始新对话，旧任务仍保留在本页历史中，可随时重新显示。')
  }, [])

  const clearExecutionRecords = useCallback(() => {
    setExecutionRecords([])
    setActiveDiffRecord(null)
    setConversationStartId(executionIdRef.current)
    setVisibleRecordLimit(DEFAULT_VISIBLE_RECORD_LIMIT)
    setShowArchivedRecords(false)
    setStatus('已清空本页 AI 面板记录。')
  }, [])

  const setRangeBindingState = useCallback((nextBinding: RangeBindingState | null) => {
    rangeBindingRef.current = nextBinding
    setRangeBinding(nextBinding)
  }, [])

  const focusCommandAt = useCallback((position: number) => {
    window.requestAnimationFrame(() => {
      const input = textareaRef.current
      if (!input) return
      input.focus()
      input.setSelectionRange(position, position)
    })
  }, [])

  const stopRangeBinding = useCallback(() => {
    setRangeBindingState(null)
  }, [setRangeBindingState])

  const closeAssistantPanel = useCallback(() => {
    setAssistantOpen(false)
    setSheetPicker(null)
    setRangeBindingState(null)
  }, [setRangeBindingState])

  const replaceRangeBindingValue = useCallback((reference: string) => {
    const binding = rangeBindingRef.current
    if (!binding || binding.value === reference) return

    const current = cmdRef.current
    const nextCmd = `${current.slice(0, binding.start)}${reference}${current.slice(binding.end)}`
    const nextBinding = {
      start: binding.start,
      end: binding.start + reference.length,
      value: reference,
    }

    cmdRef.current = nextCmd
    setCmd(nextCmd)
    setRangeBindingState(nextBinding)
  }, [setRangeBindingState])

  const startRangeBinding = useCallback((tokenStart: number, tokenEnd: number, sourceValue: string) => {
    const reference = getActiveSelectionReference() ?? '#'
    const nextCmd = `${sourceValue.slice(0, tokenStart)}${reference}${sourceValue.slice(tokenEnd)}`
    const nextBinding = {
      start: tokenStart,
      end: tokenStart + reference.length,
      value: reference,
    }

    cmdRef.current = nextCmd
    setCmd(nextCmd)
    setSheetPicker(null)
    setRangeBindingState(nextBinding)
    focusCommandAt(nextBinding.end)
  }, [focusCommandAt, setRangeBindingState])

  const insertSheetName = useCallback((sheetName: string) => {
    if (!sheetPicker) return

    const current = cmdRef.current
    const nextCmd = `${current.slice(0, sheetPicker.tokenStart)}${sheetName}${current.slice(sheetPicker.tokenEnd)}`
    const nextCursor = sheetPicker.tokenStart + sheetName.length

    cmdRef.current = nextCmd
    setCmd(nextCmd)
    setSheetPicker(null)
    focusCommandAt(nextCursor)
  }, [focusCommandAt, sheetPicker])

  function reconcileRangeBindingAfterInput(nextValue: string) {
    const binding = rangeBindingRef.current
    if (!binding) return

    if (nextValue.slice(binding.start, binding.end) === binding.value) return

    const nextStart = binding.value ? nextValue.indexOf(binding.value) : -1
    if (nextStart >= 0) {
      setRangeBindingState({
        start: nextStart,
        end: nextStart + binding.value.length,
        value: binding.value,
      })
      return
    }

    setRangeBindingState(null)
  }

  function reconcileSheetPickerAfterInput(nextValue: string, cursorPosition: number): boolean {
    if (!sheetPicker) return false
    if (cursorPosition <= sheetPicker.tokenStart || nextValue[sheetPicker.tokenStart] !== '@') {
      setSheetPicker(null)
      return false
    }

    const query = nextValue.slice(sheetPicker.tokenStart + 1, cursorPosition)
    if (/\s/.test(query) || query.includes('#')) {
      setSheetPicker(null)
      return false
    }

    setSheetPicker({
      tokenStart: sheetPicker.tokenStart,
      tokenEnd: cursorPosition,
      query,
    })
    return true
  }

  const applySettings = useCallback((nextSettings: BackendAISettings, successMessage?: string) => {
    setSettings(nextSettings)

    const activeModel = getActiveModel(nextSettings)
    if (!activeModel) {
      resetAgent()
      setConfigured(false)
      setStatus(successMessage ?? '尚未配置 AI 模型')
      return
    }

    syncFormFromModel(activeModel)
    initializeAgent(activeModel)
    setConfigured(true)
    setStatus(successMessage ? `${successMessage}，当前模型：${activeModel.model}` : `当前模型：${activeModel.model}`)
  }, [getActiveModel, initializeAgent, resetAgent, syncFormFromModel])

  useEffect(() => {
    logEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [executionRecords])

  useEffect(() => {
    cmdRef.current = cmd
  }, [cmd])

  useEffect(() => {
    if (!rangeBinding) return

    const intervalId = window.setInterval(() => {
      const reference = getActiveSelectionReference()
      if (reference) replaceRangeBindingValue(reference)
    }, 250)

    return () => window.clearInterval(intervalId)
  }, [rangeBinding, replaceRangeBindingValue])

  useEffect(() => {
    function handleEscapeKey(event: globalThis.KeyboardEvent) {
      if (event.key !== 'Escape') return

      if (sheetPicker) {
        event.preventDefault()
        setSheetPicker(null)
        return
      }

      if (rangeBinding) {
        event.preventDefault()
        stopRangeBinding()
        return
      }

      if (activeDiffRecord) {
        event.preventDefault()
        setActiveDiffRecord(null)
        return
      }

      if (configOpen) {
        event.preventDefault()
        setConfigOpen(false)
        return
      }

      if (assistantOpen) {
        event.preventDefault()
        closeAssistantPanel()
      }
    }

    document.addEventListener('keydown', handleEscapeKey)
    return () => document.removeEventListener('keydown', handleEscapeKey)
  }, [activeDiffRecord, assistantOpen, closeAssistantPanel, configOpen, rangeBinding, sheetPicker, stopRangeBinding])

  useEffect(() => {
    void (async () => {
      try {
        const nextSettings = await loadBackendAISettings()
        applySettings(nextSettings, 'AI 配置已加载')
      } catch (error) {
        setStatus(error instanceof Error ? error.message : '加载 AI 配置失败')
      } finally {
        setLoadingConfig(false)
      }
    })()

    return () => {
      agentRef.current?.cancel()
      destroyAgent()
    }
  }, [applySettings])

  async function handleDiscoverModels() {
    if (provider === 'custom' && !baseURL.trim()) return

    try {
      setDiscoveringModels(true)
      const result = await discoverBackendAIModels({
        provider,
        apiKey: apiKey.trim(),
        baseURL: baseURL.trim(),
      })

      setBaseURL(result.baseURL)
      setDiscoveredModels(result.models)
      setSelectedModels(result.models)
      setStatus(result.models.length > 0 ? `已发现 ${result.models.length} 个模型` : '没有发现可用模型')
    } catch (error) {
      setStatus(error instanceof Error ? error.message : '获取模型列表失败')
    } finally {
      setDiscoveringModels(false)
    }
  }

  function handleToggleDiscoveredModel(modelName: string) {
    setSelectedModels((prev) =>
      prev.includes(modelName)
        ? prev.filter((item) => item !== modelName)
        : [...prev, modelName],
    )
  }

  function handleAddManualModel() {
    const nextModel = manualModel.trim()
    if (!nextModel) return

    setDiscoveredModels((prev) => (prev.includes(nextModel) ? prev : [...prev, nextModel]))
    setSelectedModels((prev) => (prev.includes(nextModel) ? prev : [...prev, nextModel]))
    setManualModel('')
  }

  async function handleSaveModels() {
    if (selectedModels.length === 0) return

    try {
      setSavingConfig(true)
      const nextSettings = await saveBackendAIModels({
        provider,
        apiKey: apiKey.trim(),
        baseURL: baseURL.trim(),
        models: selectedModels,
      })

      applySettings(nextSettings, `已加入 ${selectedModels.length} 个模型`)
      setSelectedModels([])
    } catch (error) {
      setStatus(error instanceof Error ? error.message : '保存模型失败')
    } finally {
      setSavingConfig(false)
    }
  }

  async function handleSelectActiveModel(modelId: string) {
    if (!modelId || modelId === settings.activeModelId) return

    try {
      setSwitchingModel(true)
      const nextSettings = await selectActiveBackendAIModel(modelId)
      applySettings(nextSettings, '已切换模型')
    } catch (error) {
      setStatus(error instanceof Error ? error.message : '切换模型失败')
    } finally {
      setSwitchingModel(false)
    }
  }

  async function handleDeleteModel(modelId: string) {
    try {
      setSavingConfig(true)
      const nextSettings = await deleteBackendAIModel(modelId)
      applySettings(nextSettings, '模型已删除')
    } catch (error) {
      setStatus(error instanceof Error ? error.message : '删除模型失败')
    } finally {
      setSavingConfig(false)
    }
  }

  async function handleClearConfig() {
    try {
      setSavingConfig(true)
      await clearBackendAIConfig()
      resetAgent()
      setSettings(EMPTY_SETTINGS)
      setDiscoveredModels([])
      setSelectedModels([])
      setApiKey('')
      setBaseURL('')
      setManualModel('')
      setConfigured(false)
      setExecutionRecords([])
      setActiveDiffRecord(null)
      setConversationStartId(executionIdRef.current)
      setVisibleRecordLimit(DEFAULT_VISIBLE_RECORD_LIMIT)
      setShowArchivedRecords(false)
      setStatus('已清空所有 AI 模型配置')
    } catch (error) {
      setStatus(error instanceof Error ? error.message : '清空 AI 配置失败')
    } finally {
      setSavingConfig(false)
    }
  }

  function handleCommandChange(event: ChangeEvent<HTMLTextAreaElement>) {
    const nextValue = event.target.value
    const cursorPosition = event.target.selectionStart ?? nextValue.length
    const triggerIndex = cursorPosition - 1
    const trigger = triggerIndex >= 0 ? nextValue[triggerIndex] : ''

    cmdRef.current = nextValue
    setCmd(nextValue)
    reconcileRangeBindingAfterInput(nextValue)

    if (trigger === '@' && isStandaloneTrigger(nextValue, triggerIndex)) {
      setSheetNames(getWorkbookSheetNames())
      setSheetPicker({
        tokenStart: triggerIndex,
        tokenEnd: cursorPosition,
        query: '',
      })
      return
    }

    if (trigger === '#' && isStandaloneTrigger(nextValue, triggerIndex)) {
      startRangeBinding(triggerIndex, cursorPosition, nextValue)
      return
    }

    reconcileSheetPickerAfterInput(nextValue, cursorPosition)
  }

  async function handleExecute() {
    const agent = agentRef.current ?? getAgent()
    if (!agent || !cmd.trim() || running || applyingPlan) return

    setSheetPicker(null)
    stopRangeBinding()
    const instruction = cmd.trim()
    const recordId = executionIdRef.current++
    const startedAt = getCurrentTimestamp()
    const snapshot = agent.createWorkbookSnapshot(`Before: ${instruction}`)
    const nextRecord: ExecutionRecord = {
      id: recordId,
      instruction,
      status: 'running',
      message: '正在分析表格并执行指令...',
      startedAt,
      tools: [],
      plans: [],
      applyResult: null,
      verification: { passed: null, text: null },
      snapshot,
      afterSnapshot: null,
      diff: null,
      undoSteps: 0,
      undoState: 'idle',
    }
    setRunning(true)
    setVisibleRecordLimit(DEFAULT_VISIBLE_RECORD_LIMIT)
    setShowArchivedRecords(false)
    setExecutionRecords((prev) => [nextRecord, ...prev].slice(0, MAX_EXECUTION_RECORDS))

    try {
      const result = await agent.executeDetailed(instruction)
      let appliedResult: AgentApplyPlanResult | null = null

      if (result.plans.length > 0) {
        const planToApply = result.plans.at(-1)
        if (planToApply) {
          setStatus('已生成执行计划，正在自动应用...')
          appliedResult = await applyPlan(planToApply)
        }
      } else {
        setStatus(result.message || '执行完成')
      }

      const verificationSummary = appliedResult ? getVerificationSummary(appliedResult.result) : { passed: null, text: null }
      const undoSteps = getUndoStepCount(result, appliedResult)
      const afterSnapshot = agent.createWorkbookSnapshot(`After: ${instruction}`)
      const diff = agent.diffWorkbookSnapshots(snapshot.snapshotId, afterSnapshot.snapshotId, 5000)
      updateExecutionRecord(recordId, {
        status: verificationSummary.passed === false ? 'warning' : 'success',
        message: result.message || '执行完成',
        finishedAt: getCurrentTimestamp(),
        tools: result.tools,
        plans: result.plans,
        applyResult: appliedResult,
        verification: verificationSummary,
        afterSnapshot,
        diff,
        undoSteps,
      })
      setCmd('')
    } catch (error) {
      const message = error instanceof Error ? error.message : '执行失败'
      setStatus(message)
      updateExecutionRecord(recordId, {
        status: 'error',
        message,
        error: message,
        finishedAt: getCurrentTimestamp(),
      })
    } finally {
      setRunning(false)
    }
  }

  async function applyPlan(plan: AgentPlanPreview): Promise<AgentApplyPlanResult | null> {
    const agent = agentRef.current ?? getAgent()
    if (!agent || applyingPlan) return null

    try {
      setApplyingPlan(true)
      const result = await agent.applyPlan(plan.planId)
      const verificationSummary = getVerificationSummary(result.result)
      setStatus(verificationSummary.passed === false ? '计划已自动应用，但写后核对未通过' : '计划已自动应用并完成核对')
      return result
    } catch (error) {
      setStatus(error instanceof Error ? error.message : '自动应用计划失败')
      return null
    } finally {
      setApplyingPlan(false)
    }
  }

  async function handleUndoRecord(record: ExecutionRecord) {
    if (!record.snapshot || record.undoState === 'undoing' || record.undoState === 'undone') return

    const agent = agentRef.current ?? getAgent()
    if (!agent) {
      updateExecutionRecord(record.id, { undoState: 'failed', message: 'AI Agent 未初始化，无法恢复快照' })
      return
    }

    try {
      updateExecutionRecord(record.id, { undoState: 'undoing' })
      const result = agent.restoreWorkbookSnapshot(record.snapshot.snapshotId)
      if (!result.success) {
        updateExecutionRecord(record.id, { undoState: 'failed', message: result.message })
        setStatus(result.message)
        return
      }

      updateExecutionRecord(record.id, { status: 'undone', undoState: 'undone', message: result.message })
      setStatus(`已回到「${record.instruction}」执行前`)
    } catch (error) {
      const message = error instanceof Error ? error.message : '撤销失败'
      updateExecutionRecord(record.id, { undoState: 'failed', message })
      setStatus(message)
    }
  }

  function handleReapplyRecord(record: ExecutionRecord) {
    if (!record.afterSnapshot || record.undoState === 'undoing') return

    const agent = agentRef.current ?? getAgent()
    if (!agent) {
      updateExecutionRecord(record.id, { undoState: 'failed', message: 'AI Agent 未初始化，无法再次应用' })
      return
    }

    const result = agent.restoreWorkbookSnapshot(record.afterSnapshot.snapshotId)
    if (!result.success) {
      updateExecutionRecord(record.id, { undoState: 'failed', message: result.message })
      setStatus(result.message)
      return
    }

    updateExecutionRecord(record.id, {
      status: record.verification.passed === false ? 'warning' : 'success',
      undoState: 'idle',
      message: `已再次应用本次 AI 操作：${record.afterSnapshot.sheetCount} 个工作表，${record.afterSnapshot.cellCount} 个单元格。`,
    })
    setStatus(`已再次应用「${record.instruction}」`)
  }

  function handleCancel() {
    const agent = agentRef.current ?? getAgent()
    if (!agent || !running) return
    agent.cancel('User canceled execution')
  }

  function handleKeyDown(event: KeyboardEvent<HTMLTextAreaElement>) {
    if (event.key === 'Escape') {
      if (sheetPicker) {
        event.preventDefault()
        event.stopPropagation()
        setSheetPicker(null)
        return
      }

      if (rangeBinding) {
        event.preventDefault()
        event.stopPropagation()
        stopRangeBinding()
        return
      }
    }

    if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
      event.preventDefault()
      void handleExecute()
    }
  }

  const activeModel = getActiveModel(settings)
  const canDiscoverModels = provider !== 'custom' || !!baseURL.trim()
  const canSaveModels = canDiscoverModels && selectedModels.length > 0
  const visibleSheetNames = sheetPicker
    ? sheetNames.filter((sheetName) => sheetName.toLowerCase().includes(sheetPicker.query.toLowerCase()))
    : []
  const currentConversationRecords = executionRecords.filter((record) => record.id >= conversationStartId)
  const archivedRecordCount = executionRecords.length - currentConversationRecords.length
  const displayableRecords = showArchivedRecords ? executionRecords : currentConversationRecords
  const visibleExecutionRecords = displayableRecords.slice(0, visibleRecordLimit)
  const hiddenRecordCount = Math.max(0, displayableRecords.length - visibleExecutionRecords.length)

  return (
    <>
      <div style={anchorStyle}>
        <button
          style={{
            ...toolbarButtonStyle,
            borderColor: assistantOpen ? '#2563eb' : 'rgba(148, 163, 184, 0.24)',
            color: '#0f172a',
          }}
          onClick={() => {
            if (assistantOpen) {
              closeAssistantPanel()
              return
            }
            setAssistantOpen(true)
          }}
        >
          <span style={toolbarDotStyle(configured, running || applyingPlan)} />
          <span>AI 助手</span>
        </button>

        {canConfigureAI ? (
          <button
            style={{
              ...toolbarButtonStyle,
              borderColor: configOpen ? '#2563eb' : 'rgba(148, 163, 184, 0.24)',
              color: configured ? '#0f172a' : '#475569',
            }}
            onClick={() => setConfigOpen(true)}
          >
            <span style={toolbarDotStyle(configured, false)} />
            <span>AI 配置</span>
          </button>
        ) : null}

        {assistantOpen && (
          <div style={assistantPanelStyle}>
            <div style={panelHeaderStyle}>
              <div>
                <div style={panelTitleStyle}>AI 助手</div>
                <div style={panelSubtitleStyle}>
                  {configured ? '执行、写入和核对会自动完成，面板只展示过程。' : '请先在 AI 配置里加入模型。'}
                </div>
              </div>
              <button style={closeButtonStyle} onClick={closeAssistantPanel}>
                关闭
              </button>
            </div>

            <div style={panelContentStyle}>
              <div style={statusCardStyle}>
                <div style={statusTitleStyle}>当前状态</div>
                <div style={panelSubtitleStyle}>{status}</div>
              </div>

              <div style={labelStyle}>当前模型</div>
              <select
                value={activeModel?.id ?? ''}
                onChange={(event) => void handleSelectActiveModel(event.target.value)}
                style={selectStyle}
                disabled={!configured || switchingModel || running || applyingPlan}
              >
                {settings.models.length === 0 ? (
                  <option value="">暂无模型</option>
                ) : (
                  settings.models.map((model) => (
                    <option key={model.id} value={model.id}>
                      {model.model} / {PROVIDER_LABELS[(VALID_PROVIDERS.includes(model.provider as LLMProvider)
                        ? model.provider
                        : 'custom') as LLMProvider]}
                    </option>
                  ))
                )}
              </select>

              <div style={labelStyle}>指令</div>
              <textarea
                ref={textareaRef}
                value={cmd}
                onChange={handleCommandChange}
                onKeyDown={handleKeyDown}
                placeholder="输入自然语言指令，例如：按 PO# 匹配两张表，把目标表的 Unit Price 按订单总额 / 数量更新"
                style={textareaStyle}
                disabled={!configured || running || applyingPlan}
              />

              {sheetPicker ? (
                <div style={sheetPickerStyle}>
                  <div style={sheetPickerHeaderStyle}>选择工作表</div>
                  {visibleSheetNames.length > 0 ? (
                    visibleSheetNames.map((sheetName) => (
                      <button
                        key={sheetName}
                        type="button"
                        style={sheetOptionStyle}
                        onMouseDown={(event) => {
                          event.preventDefault()
                          insertSheetName(sheetName)
                        }}
                      >
                        {sheetName}
                      </button>
                    ))
                  ) : (
                    <div style={sheetPickerEmptyStyle}>未找到工作表</div>
                  )}
                </div>
              ) : null}

              {rangeBinding ? (
                <div style={rangeBindingBarStyle}>
                  <span>
                    正在监听表格选区：{rangeBinding.value === '#' ? '请选择单元格或区域' : rangeBinding.value}
                  </span>
                  <button type="button" style={inlineTextButtonStyle} onClick={stopRangeBinding}>
                    停止监听
                  </button>
                </div>
              ) : (
                <div style={helperTextStyle}>提示：@ 插入工作表名，# 绑定当前选区并实时更新到指令框。</div>
              )}

              <div style={commandRowStyle}>
                <button
                  style={{ ...primaryButtonStyle, minWidth: 96 }}
                  onClick={() => void (running ? handleCancel() : handleExecute())}
                  disabled={!configured || applyingPlan || (!running && !cmd.trim())}
                >
                  {running ? '停止' : '执行'}
                </button>
                <button
                  style={secondaryButtonStyle}
                  onClick={startNewConversation}
                  disabled={running || applyingPlan || currentConversationRecords.length === 0}
                >
                  新对话
                </button>
                <button
                  style={secondaryButtonStyle}
                  onClick={clearExecutionRecords}
                  disabled={running || applyingPlan || executionRecords.length === 0}
                >
                  清空本页记录
                </button>
              </div>

              {!configured ? (
                <div style={hintCardStyle}>
                  先在 AI 配置里通过 `/v1/models` 获取模型列表，把需要的模型加入系统，再回到这里执行。
                </div>
              ) : null}

              <div style={historyAreaStyle}>
                {executionRecords.length === 0 || displayableRecords.length === 0 ? (
                  <div style={emptyHistoryStyle}>
                    {executionRecords.length === 0
                      ? '执行结果会以任务卡片展示：计划、操作步骤、写入位置、核对结果和撤销入口会分开显示。'
                      : '当前是新对话，旧任务已隐藏；需要回退旧操作时可以显示本页历史。'}
                  </div>
                ) : (
                  visibleExecutionRecords.map((record, index) => (
                    <ExecutionRecordCard
                      key={record.id}
                      record={record}
                      applying={applyingPlan && record.status === 'running'}
                      initiallyExpanded={index === 0}
                      onUndo={() => void handleUndoRecord(record)}
                      onReapply={() => handleReapplyRecord(record)}
                      onOpenDiff={() => setActiveDiffRecord(record)}
                    />
                  ))
                )}
                {hiddenRecordCount > 0 ? (
                  <button
                    type="button"
                    style={loadMoreButtonStyle}
                    onClick={() => setVisibleRecordLimit((prev) => prev + VISIBLE_RECORD_INCREMENT)}
                  >
                    显示更早 {Math.min(hiddenRecordCount, VISIBLE_RECORD_INCREMENT)} 条
                  </button>
                ) : null}
                {archivedRecordCount > 0 ? (
                  <button
                    type="button"
                    style={loadMoreButtonStyle}
                    onClick={() => {
                      setShowArchivedRecords((prev) => !prev)
                      setVisibleRecordLimit(DEFAULT_VISIBLE_RECORD_LIMIT)
                    }}
                  >
                    {showArchivedRecords ? '隐藏旧对话' : `显示本页历史 ${archivedRecordCount} 条`}
                  </button>
                ) : null}
                <div ref={logEndRef} />
              </div>
            </div>
          </div>
        )}
      </div>

      {configOpen && canConfigureAI ? (
        <ConfigModal
          loadingConfig={loadingConfig}
          savingConfig={savingConfig}
          discoveringModels={discoveringModels}
          provider={provider}
          apiKey={apiKey}
          baseURL={baseURL}
          manualModel={manualModel}
          settings={settings}
          discoveredModels={discoveredModels}
          selectedModels={selectedModels}
          activeModel={activeModel}
          switchingModel={switchingModel}
          canDiscoverModels={canDiscoverModels}
          canSaveModels={canSaveModels}
          setProvider={setProvider}
          setApiKey={setApiKey}
          setBaseURL={setBaseURL}
          setManualModel={setManualModel}
          onClose={() => setConfigOpen(false)}
          onDiscover={() => void handleDiscoverModels()}
          onClear={() => void handleClearConfig()}
          onSave={() => void handleSaveModels()}
          onAddManual={handleAddManualModel}
          onToggleModel={handleToggleDiscoveredModel}
          onSelectActive={(modelId) => void handleSelectActiveModel(modelId)}
          onDeleteModel={(modelId) => void handleDeleteModel(modelId)}
        />
      ) : null}

      {activeDiffRecord?.diff ? (
        <DiffDetailModal
          record={activeDiffRecord}
          diff={activeDiffRecord.diff}
          onClose={() => setActiveDiffRecord(null)}
        />
      ) : null}
    </>
  )
}

function ExecutionRecordCard({
  record,
  applying,
  initiallyExpanded,
  onUndo,
  onReapply,
  onOpenDiff,
}: {
  record: ExecutionRecord
  applying: boolean
  initiallyExpanded: boolean
  onUndo(): void
  onReapply(): void
  onOpenDiff(): void
}) {
  const [detailsOpen, setDetailsOpen] = useState(initiallyExpanded || record.status === 'running')
  const metrics = getExecutionMetrics(record)
  const latestPlan = record.plans.at(-1)
  const verificationTone = record.verification.passed === false ? 'bad' : record.verification.passed === true ? 'ok' : 'neutral'
  const canUndo = !!record.snapshot && record.undoState !== 'undone' && record.undoState !== 'undoing'
  const canReapply = !!record.afterSnapshot && record.undoState === 'undone'
  const statusLabel: Record<ExecutionStatus, string> = {
    running: applying ? '执行中' : '处理中',
    success: '已完成',
    warning: '需检查',
    error: '失败',
    undone: '已回退',
  }

  return (
    <div
      style={{
        ...executionCardStyle,
        borderColor: record.status === 'error'
          ? 'rgba(220, 38, 38, 0.28)'
          : record.status === 'warning'
            ? 'rgba(217, 119, 6, 0.30)'
            : record.status === 'undone'
              ? 'rgba(100, 116, 139, 0.24)'
              : executionCardStyle.borderColor,
      }}
    >
      <div style={executionHeaderStyle}>
        <div style={executionTitleWrapStyle}>
          <div style={executionCommandStyle}>{record.instruction}</div>
          <div style={executionMetaStyle}>
            {formatRecordTime(record.startedAt)}
            {record.finishedAt ? ` - ${formatRecordTime(record.finishedAt)}` : ''}
          </div>
        </div>
        <span
          style={{
            ...executionStatusBadgeStyle,
            background: record.status === 'error'
              ? '#fee2e2'
              : record.status === 'warning'
                ? '#fef3c7'
                : record.status === 'undone'
                  ? '#f1f5f9'
                  : record.status === 'running'
                    ? '#dbeafe'
                    : '#dcfce7',
            color: record.status === 'error'
              ? '#b91c1c'
              : record.status === 'warning'
                ? '#92400e'
                : record.status === 'undone'
                  ? '#475569'
                  : record.status === 'running'
                    ? '#1d4ed8'
                    : '#166534',
          }}
        >
          {statusLabel[record.status]}
        </span>
      </div>

      <div style={metricGridStyle}>
        {metrics.map((metric) => (
          <div
            key={metric.label}
            style={{
              ...metricTileStyle,
              color: metric.tone === 'bad' ? '#b91c1c' : metric.tone === 'warn' ? '#92400e' : '#334155',
            }}
          >
            <span style={metricValueStyle}>{metric.value}</span>
            <span>{metric.label}</span>
          </div>
        ))}
      </div>

      <div style={recordSummaryStyle}>
        <span>{record.message}</span>
        <button type="button" style={inlineTextButtonStyle} onClick={() => setDetailsOpen((prev) => !prev)}>
          {detailsOpen ? '收起详情' : '展开详情'}
        </button>
      </div>

      {detailsOpen ? (
        <>
          {latestPlan ? <PlanPreviewCard plan={latestPlan} applying={record.status === 'running'} compact /> : null}

          {record.diff ? <DiffSummaryCard diff={record.diff} onOpenFull={onOpenDiff} /> : null}

          {record.tools.length > 0 ? (
            <div style={operationTableStyle}>
              <div style={operationTableHeaderStyle}>
                <span>类型</span>
                <span>操作</span>
                <span>范围</span>
                <span>结果</span>
              </div>
              {record.tools.slice(0, 12).map((tool, index) => {
                const state = getToolResultState(tool)
                const result = asRecord(tool.result)
                return (
                  <div key={`${record.id}-${tool.toolName}-${index}`} style={operationRowStyle}>
                    <span style={operationBadgeStyle}>{getToolBadge(tool.toolName)}</span>
                    <span style={operationNameStyle}>{getToolTitle(tool.toolName)}</span>
                    <span style={operationScopeStyle}>{getToolScope(tool)}</span>
                    <span
                      style={{
                        ...operationStateStyle,
                        color: state === 'error' ? '#b91c1c' : state === 'warn' ? '#92400e' : state === 'ok' ? '#166534' : '#475569',
                      }}
                    >
                      {state === 'error' ? asString(result?.error) ?? '失败' : state === 'warn' ? '已应用，需注意' : state === 'ok' ? '已应用' : '已读取'}
                    </span>
                  </div>
                )
              })}
              {record.tools.length > 12 ? <div style={tableFooterStyle}>还有 {record.tools.length - 12} 个步骤未展开显示。</div> : null}
            </div>
          ) : null}

          {record.verification.text ? (
            <div
              style={{
                ...verificationCardStyle,
                borderColor: verificationTone === 'bad' ? 'rgba(220, 38, 38, 0.24)' : verificationTone === 'ok' ? 'rgba(22, 163, 74, 0.22)' : 'rgba(148, 163, 184, 0.18)',
                background: verificationTone === 'bad' ? '#fff5f5' : verificationTone === 'ok' ? '#f0fdf4' : '#f8fafc',
              }}
            >
              <div style={summaryTitleStyle}>写后核对</div>
              <div style={summaryTextStyle}>{record.verification.text}</div>
            </div>
          ) : null}

          {record.error ? <div style={errorTextStyle}>{record.error}</div> : null}
        </>
      ) : null}

      <div style={recordActionsStyle}>
        <div style={recordButtonGroupStyle}>
          <button
            style={canUndo ? secondaryButtonStyle : disabledButtonStyle}
            onClick={onUndo}
            disabled={!canUndo}
            title={record.snapshot ? `恢复执行前快照：${record.snapshot.sheetCount} 个工作表，${record.snapshot.cellCount} 个单元格` : '本次没有可恢复的快照'}
          >
            {record.undoState === 'undoing' ? '正在回退...' : record.undoState === 'undone' ? '已回到执行前' : '回到此次更新前'}
          </button>
          <button
            style={canReapply ? primaryButtonStyle : disabledButtonStyle}
            onClick={onReapply}
            disabled={!canReapply}
            title={record.afterSnapshot ? `恢复执行后快照：${record.afterSnapshot.sheetCount} 个工作表，${record.afterSnapshot.cellCount} 个单元格` : '本次没有执行后快照'}
          >
            再次应用
          </button>
        </div>
        <span style={recordHintStyle}>
          {record.snapshot && record.afterSnapshot
            ? `前后快照：${record.snapshot.sheetCount} 表 / ${record.snapshot.cellCount} 单元格。`
            : record.snapshot
              ? `执行前快照：${record.snapshot.sheetCount} 表 / ${record.snapshot.cellCount} 单元格。`
              : '本次没有可恢复快照。'}
        </span>
      </div>
    </div>
  )
}

function DiffSummaryCard({
  diff,
  onOpenFull,
}: {
  diff: AgentSnapshotDiff
  onOpenFull(): void
}) {
  if (!diff.success) {
    return (
      <div style={{ ...diffCardStyle, borderColor: 'rgba(220, 38, 38, 0.24)', background: '#fff5f5' }}>
        <div style={diffHeaderStyle}>
          <div style={summaryTitleStyle}>变更对比</div>
          <button style={compactButtonStyle} onClick={onOpenFull}>查看详情</button>
        </div>
        <div style={summaryTextStyle}>{diff.message}</div>
      </div>
    )
  }

  const displayChanges = getDisplayableDiffChanges(diff)

  return (
    <div style={diffCardStyle}>
      <div style={diffHeaderStyle}>
        <div>
          <div style={summaryTitleStyle}>变更对比</div>
          <div style={diffSubtitleStyle}>
            值 {diff.valueChangeCount}，公式 {diff.formulaChangeCount}，格式 {diff.formatChangeCount}
          </div>
        </div>
        <span style={diffBadgeStyle}>{diff.changedCellCount} 项变化</span>
      </div>

      {diff.addedSheetNames.length > 0 || diff.removedSheetNames.length > 0 ? (
        <div style={diffSheetNoticeStyle}>
          {diff.addedSheetNames.length > 0 ? `新增工作表：${diff.addedSheetNames.join(', ')}` : ''}
          {diff.addedSheetNames.length > 0 && diff.removedSheetNames.length > 0 ? '；' : ''}
          {diff.removedSheetNames.length > 0 ? `移除工作表：${diff.removedSheetNames.join(', ')}` : ''}
        </div>
      ) : null}

      {displayChanges.length === 0 ? (
        <div style={diffEmptyStyle}>
          {diff.formatChangeCount > 0 ? '本次只有格式变化，当前不展示格式明细。' : '没有检测到单元格值或公式变化。'}
        </div>
      ) : (
        <div style={diffTableStyle}>
          <div style={diffTableHeaderStyle}>
            <span>位置</span>
            <span>类型</span>
            <span>执行前</span>
            <span>执行后</span>
          </div>
          {displayChanges.slice(0, 12).map((change, index) => (
            <div key={`${change.sheetName}-${change.cell}-${index}`} style={diffRowStyle}>
              <span style={diffCellRefStyle}>{change.sheetName}!{change.cell}</span>
              <span style={diffKindStyle}>{getDiffKindLabel(change.kind)}</span>
              <span style={diffValueStyle}>{formatDiffCellValue(change.beforeValue, change.beforeFormula)}</span>
              <span style={diffValueStyle}>{formatDiffCellValue(change.afterValue, change.afterFormula)}</span>
            </div>
          ))}
        </div>
      )}

      {diff.truncated || displayChanges.length > 12 ? (
        <div style={diffFooterStyle}>仅展示前 {Math.min(displayChanges.length, 12)} 条可解释差异，完整差异已用于快照回退和再次应用。</div>
      ) : null}

      <div style={diffActionsStyle}>
        <button style={secondaryButtonStyle} onClick={onOpenFull}>
          查看全部差异
        </button>
      </div>
    </div>
  )
}

function DiffDetailModal({
  record,
  diff,
  onClose,
}: {
  record: ExecutionRecord
  diff: AgentSnapshotDiff
  onClose(): void
}) {
  const [filter, setFilter] = useState<DiffFilter>('all')
  const displayChanges = getDisplayableDiffChanges(diff)
  const filteredChanges = displayChanges.filter((change) => {
    if (filter === 'all') return true
    if (filter === 'value') return change.kind === 'value' || change.kind === 'value_formula'
    if (filter === 'formula') return change.kind === 'formula' || change.kind === 'value_formula'
    return false
  })

  const filterOptions: Array<{ value: DiffFilter; label: string; count: number }> = [
    { value: 'all', label: '全部', count: displayChanges.length },
    { value: 'value', label: '值', count: displayChanges.filter((change) => change.kind === 'value' || change.kind === 'value_formula').length },
    { value: 'formula', label: '公式', count: displayChanges.filter((change) => change.kind === 'formula' || change.kind === 'value_formula').length },
  ]

  return (
    <div style={modalBackdropStyle} onClick={onClose}>
      <div style={diffModalStyle} onClick={(event) => event.stopPropagation()}>
        <div style={panelHeaderStyle}>
          <div>
            <div style={panelTitleStyle}>完整差异</div>
            <div style={panelSubtitleStyle}>{record.instruction}</div>
          </div>
          <button style={closeButtonStyle} onClick={onClose}>
            关闭
          </button>
        </div>

        {!diff.success ? (
          <div style={{ ...errorTextStyle, marginTop: 14 }}>{diff.message}</div>
        ) : (
          <>
            <div style={diffOverviewGridStyle}>
              <div style={diffOverviewTileStyle}>
                <span style={metricValueStyle}>{diff.changedCellCount}</span>
                <span>变化单元格</span>
              </div>
              <div style={diffOverviewTileStyle}>
                <span style={metricValueStyle}>{diff.valueChangeCount}</span>
                <span>值变化</span>
              </div>
              <div style={diffOverviewTileStyle}>
                <span style={metricValueStyle}>{diff.formulaChangeCount}</span>
                <span>公式变化</span>
              </div>
              <div style={diffOverviewTileStyle}>
                <span style={metricValueStyle}>{diff.formatChangeCount}</span>
                <span>格式变化</span>
              </div>
            </div>

            <div style={diffToolbarStyle}>
              <div style={providerRowStyle}>
                {filterOptions.map((option) => (
                  <button
                    key={option.value}
                    style={{
                      ...tagButtonStyle,
                      background: filter === option.value ? '#0369a1' : '#e0f2fe',
                      color: filter === option.value ? '#f0f9ff' : '#0369a1',
                    }}
                    onClick={() => setFilter(option.value)}
                  >
                    {option.label} ({option.count})
                  </button>
                ))}
              </div>
              <div style={diffFooterStyle}>
                {diff.formatChangeCount > 0 ? `格式变化 ${diff.formatChangeCount} 项只保留统计，不展示明细。` : ''}
                {diff.formatChangeCount > 0 ? ' ' : ''}
                {diff.truncated ? `差异过多，已保留前 ${displayChanges.length} 条可解释差异。` : `已加载 ${displayChanges.length} 条可解释差异。`}
              </div>
            </div>

            {diff.addedSheetNames.length > 0 || diff.removedSheetNames.length > 0 ? (
              <div style={diffSheetNoticeStyle}>
                {diff.addedSheetNames.length > 0 ? `新增工作表：${diff.addedSheetNames.join(', ')}` : ''}
                {diff.addedSheetNames.length > 0 && diff.removedSheetNames.length > 0 ? '；' : ''}
                {diff.removedSheetNames.length > 0 ? `移除工作表：${diff.removedSheetNames.join(', ')}` : ''}
              </div>
            ) : null}

            <div style={fullDiffTableStyle}>
              <div style={fullDiffTableHeaderStyle}>
                <span>位置</span>
                <span>类型</span>
                <span>执行前</span>
                <span>执行后</span>
              </div>
              {filteredChanges.length === 0 ? (
                <div style={fullDiffEmptyStyle}>
                  {diff.formatChangeCount > 0 && displayChanges.length === 0
                    ? '本次只有格式变化，当前不展示格式明细。'
                    : '当前筛选条件下没有值或公式差异。'}
                </div>
              ) : (
                filteredChanges.map((change, index) => (
                  <div key={`${change.sheetName}-${change.cell}-${change.kind}-${index}`} style={fullDiffRowStyle}>
                    <span style={diffCellRefStyle}>{change.sheetName}!{change.cell}</span>
                    <span style={diffKindStyle}>{getDiffKindLabel(change.kind)}</span>
                    <span style={fullDiffValueStyle}>{formatDiffCellValue(change.beforeValue, change.beforeFormula)}</span>
                    <span style={fullDiffValueStyle}>{formatDiffCellValue(change.afterValue, change.afterFormula)}</span>
                  </div>
                ))
              )}
            </div>
          </>
        )}
      </div>
    </div>
  )
}

function PlanPreviewCard({
  plan,
  applying,
  compact = false,
}: {
  plan: AgentPlanPreview
  applying: boolean
  compact?: boolean
}) {
  const preview = getPlanPreview(plan)
  const stats = getPreviewStats(preview)
  const matchRows = getMatchRows(preview)
  const matchAnalysisRows = getMatchAnalysisRows(preview)
  const derivedRows = getDerivedRows(preview)
  const writeRows = getWriteRows(preview)
  const issueRows = getIssueRows(preview)
  const hasBlockingIssues = issueRows.length > 0

  return (
    <div
      style={{
        ...planCardStyle,
        maxHeight: compact ? 260 : planCardStyle.maxHeight,
        borderColor: hasBlockingIssues ? 'rgba(217, 119, 6, 0.32)' : planCardStyle.borderColor,
        background: hasBlockingIssues ? '#fffbeb' : planCardStyle.background,
      }}
    >
      <div style={planCardHeaderStyle}>
        <div>
          <div style={planTitleStyle}>{applying ? '正在执行计划' : '最近执行计划'}</div>
          <div style={planMetaStyle}>{plan.planId}</div>
        </div>
        <span style={planStateBadgeStyle}>{applying ? '自动应用中' : '已自动应用'}</span>
      </div>

      <div style={planBodyStyle}>{summarizePlanPreview(plan)}</div>

      {stats.length > 0 ? (
        <div style={statGridStyle}>
          {stats.map((stat) => (
            <div
              key={stat.label}
              style={{
                ...statPillStyle,
                borderColor: stat.tone === 'bad' ? 'rgba(220, 38, 38, 0.24)' : stat.tone === 'warn' ? 'rgba(217, 119, 6, 0.24)' : 'rgba(37, 99, 235, 0.18)',
                color: stat.tone === 'bad' ? '#dc2626' : stat.tone === 'warn' ? '#b45309' : '#1d4ed8',
              }}
            >
              <span style={statValueStyle}>{stat.value}</span>
              <span>{stat.label}</span>
            </div>
          ))}
        </div>
      ) : null}

      <DetailList title="匹配字段" emptyText="没有匹配字段信息" items={matchRows} />
      <DetailList title="匹配证据" emptyText="没有共享值证据" items={matchAnalysisRows} />
      <DetailList title="推导字段" emptyText="没有派生计算" items={derivedRows} />
      <DetailList title="写入位置" emptyText="没有明确写入位置" items={writeRows} />
      {issueRows.length > 0 ? <DetailList title="需要注意" items={issueRows} tone="bad" /> : null}
    </div>
  )
}

function DetailList({
  title,
  items,
  emptyText,
  tone,
}: {
  title: string
  items: string[]
  emptyText?: string
  tone?: 'bad'
}) {
  const visibleItems = items.length > 0 ? items : emptyText ? [emptyText] : []
  if (visibleItems.length === 0) return null

  return (
    <div style={detailBlockStyle}>
      <div style={detailTitleStyle}>{title}</div>
      <div style={detailListStyle}>
        {visibleItems.map((item, index) => (
          <div
            key={`${title}-${index}`}
            style={{
              ...detailItemStyle,
              color: tone === 'bad' ? '#b91c1c' : '#334155',
              background: tone === 'bad' ? '#fff5f5' : '#ffffff',
            }}
          >
            {item}
          </div>
        ))}
      </div>
    </div>
  )
}

function ConfigModal(props: {
  loadingConfig: boolean
  savingConfig: boolean
  discoveringModels: boolean
  provider: LLMProvider
  apiKey: string
  baseURL: string
  manualModel: string
  settings: BackendAISettings
  discoveredModels: string[]
  selectedModels: string[]
  activeModel: BackendAIModel | null
  switchingModel: boolean
  canDiscoverModels: boolean
  canSaveModels: boolean
  setProvider(value: LLMProvider): void
  setApiKey(value: string): void
  setBaseURL(value: string): void
  setManualModel(value: string): void
  onClose(): void
  onDiscover(): void
  onClear(): void
  onSave(): void
  onAddManual(): void
  onToggleModel(modelName: string): void
  onSelectActive(modelId: string): void
  onDeleteModel(modelId: string): void
}) {
  return (
    <div style={modalBackdropStyle} onClick={props.onClose}>
      <div style={modalStyle} onClick={(event) => event.stopPropagation()}>
        <div style={panelHeaderStyle}>
          <div>
            <div style={panelTitleStyle}>AI 配置</div>
            <div style={panelSubtitleStyle}>先获取模型，再把需要的模型加入系统。</div>
          </div>
          <button style={closeButtonStyle} onClick={props.onClose}>
            关闭
          </button>
        </div>

        <div style={sectionStyle}>
          <div style={labelStyle}>Provider</div>
          <div style={providerRowStyle}>
            {VALID_PROVIDERS.map((item) => (
              <button
                key={item}
                style={{
                  ...tagButtonStyle,
                  background: props.provider === item ? '#2563eb' : '#e2e8f0',
                  color: props.provider === item ? '#eff6ff' : '#334155',
                }}
                onClick={() => props.setProvider(item)}
                disabled={props.loadingConfig || props.savingConfig || props.discoveringModels}
              >
                {PROVIDER_LABELS[item]}
              </button>
            ))}
          </div>

          <div style={labelStyle}>API Base URL</div>
          <input
            type="url"
            value={props.baseURL}
            onChange={(event) => props.setBaseURL(event.target.value)}
            placeholder={PROVIDER_BASE_URL_PLACEHOLDER[props.provider]}
            style={inputStyle}
            disabled={props.loadingConfig || props.savingConfig || props.discoveringModels}
          />

          <div style={labelStyle}>API Key</div>
          <input
            type="password"
            value={props.apiKey}
            onChange={(event) => props.setApiKey(event.target.value)}
            placeholder="输入 API Key，没有则留空"
            style={inputStyle}
            disabled={props.loadingConfig || props.savingConfig || props.discoveringModels}
          />

          <div style={helperTextStyle}>
            系统会根据当前 Base URL 请求 `/v1/models`，获取可用模型列表，再把选中的模型写入本地后端配置。
          </div>

          <div style={actionsRowStyle}>
            <button
              style={primaryButtonStyle}
              onClick={props.onDiscover}
              disabled={!props.canDiscoverModels || props.discoveringModels || props.savingConfig}
            >
              {props.discoveringModels ? '获取中...' : '获取模型列表'}
            </button>
            <button
              style={secondaryButtonStyle}
              onClick={props.onClear}
              disabled={props.loadingConfig || props.savingConfig || props.discoveringModels}
            >
              清空所有模型
            </button>
          </div>
        </div>

        <div style={dividerStyle} />

        <div style={sectionStyle}>
          <div style={sectionHeaderStyle}>
            <div>
              <div style={panelTitleStyle}>发现的模型</div>
              <div style={panelSubtitleStyle}>点击模型进行选择，再加入系统。</div>
            </div>
            <button style={primaryButtonStyle} onClick={props.onSave} disabled={!props.canSaveModels || props.savingConfig}>
              {props.savingConfig ? '保存中...' : `加入系统${props.selectedModels.length > 0 ? ` (${props.selectedModels.length})` : ''}`}
            </button>
          </div>

          <div style={commandRowStyle}>
            <input
              value={props.manualModel}
              onChange={(event) => props.setManualModel(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') {
                  event.preventDefault()
                  props.onAddManual()
                }
              }}
              placeholder="手动补充模型名，例如 gpt-4o-mini"
              style={{ ...inputStyle, flex: 1 }}
              disabled={props.savingConfig}
            />
            <button style={secondaryButtonStyle} onClick={props.onAddManual} disabled={props.savingConfig}>
              加入候选
            </button>
          </div>

          {props.discoveredModels.length === 0 ? (
            <div style={emptyStateStyle}>还没有模型列表。先点击“获取模型列表”，或者手动补充模型名。</div>
          ) : (
            <div style={chipGridStyle}>
              {props.discoveredModels.map((modelName) => {
                const selected = props.selectedModels.includes(modelName)
                return (
                  <button
                    key={modelName}
                    style={{
                      ...chipButtonStyle,
                      background: selected ? '#dbeafe' : '#f8fafc',
                      borderColor: selected ? '#60a5fa' : 'rgba(148, 163, 184, 0.22)',
                      color: selected ? '#1d4ed8' : '#334155',
                    }}
                    onClick={() => props.onToggleModel(modelName)}
                  >
                    {modelName}
                  </button>
                )
              })}
            </div>
          )}
        </div>

        <div style={dividerStyle} />

        <div style={sectionStyle}>
          <div>
            <div style={panelTitleStyle}>系统中的模型</div>
            <div style={panelSubtitleStyle}>AI 助手会从这里读取模型，并允许切换当前执行模型。</div>
          </div>

          {props.settings.models.length === 0 ? (
            <div style={emptyStateStyle}>当前还没有加入系统的模型。</div>
          ) : (
            <div style={savedModelListStyle}>
              {props.settings.models.map((model) => {
                const isActive = model.id === props.activeModel?.id
                const providerLabel = PROVIDER_LABELS[
                  (VALID_PROVIDERS.includes(model.provider as LLMProvider) ? model.provider : 'custom') as LLMProvider
                ]

                return (
                  <div key={model.id} style={savedModelCardStyle}>
                    <div style={savedModelHeaderStyle}>
                      <div>
                        <div style={savedModelTitleStyle}>{model.model}</div>
                        <div style={savedModelMetaStyle}>{providerLabel}</div>
                      </div>
                      {isActive ? <span style={activeBadgeStyle}>当前</span> : null}
                    </div>

                    <div style={savedModelMetaStyle}>{model.baseURL}</div>

                    <div style={savedModelActionsStyle}>
                      <button
                        style={secondaryButtonStyle}
                        onClick={() => props.onSelectActive(model.id)}
                        disabled={isActive || props.switchingModel}
                      >
                        设为当前
                      </button>
                      <button
                        style={dangerButtonStyle}
                        onClick={() => props.onDeleteModel(model.id)}
                        disabled={props.savingConfig}
                      >
                        删除
                      </button>
                    </div>
                  </div>
                )
              })}
            </div>
          )}
        </div>
      </div>
    </div>
  )
}

function toolbarDotStyle(configured: boolean, running: boolean): CSSProperties {
  return {
    width: 8,
    height: 8,
    borderRadius: 999,
    background: running ? '#d97706' : configured ? '#16a34a' : '#94a3b8',
    flexShrink: 0,
  }
}

const anchorStyle: CSSProperties = {
  position: 'relative',
  display: 'flex',
  alignItems: 'center',
  gap: 10,
  pointerEvents: 'auto',
}

const toolbarButtonStyle: CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  gap: 8,
  height: 34,
  padding: '0 14px',
  borderRadius: 999,
  border: '1px solid rgba(148, 163, 184, 0.24)',
  background: 'rgba(255, 255, 255, 0.96)',
  boxShadow: '0 8px 20px rgba(15, 23, 42, 0.08)',
  fontSize: 13,
  fontWeight: 600,
  cursor: 'pointer',
}

const assistantPanelStyle: CSSProperties = {
  position: 'absolute',
  top: 46,
  right: 0,
  width: 580,
  height: 'min(calc(100vh - 96px), 980px)',
  overflow: 'hidden',
  display: 'flex',
  flexDirection: 'column',
  borderRadius: 14,
  border: '1px solid rgba(148, 163, 184, 0.22)',
  background: 'rgba(255, 255, 255, 0.98)',
  boxShadow: '0 24px 70px rgba(15, 23, 42, 0.18)',
  padding: 18,
  color: '#0f172a',
  pointerEvents: 'auto',
}

const panelContentStyle: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  gap: 10,
  marginTop: 14,
  minHeight: 0,
  flex: 1,
  overflow: 'auto',
  paddingRight: 2,
}

const modalBackdropStyle: CSSProperties = {
  position: 'fixed',
  inset: 0,
  zIndex: 100,
  background: 'rgba(15, 23, 42, 0.32)',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  padding: 24,
  pointerEvents: 'auto',
}

const modalStyle: CSSProperties = {
  width: 'min(760px, calc(100vw - 48px))',
  maxHeight: 'min(86vh, 900px)',
  overflow: 'auto',
  borderRadius: 14,
  border: '1px solid rgba(148, 163, 184, 0.22)',
  background: '#ffffff',
  boxShadow: '0 30px 80px rgba(15, 23, 42, 0.28)',
  padding: 20,
  color: '#0f172a',
}

const panelHeaderStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'flex-start',
  justifyContent: 'space-between',
  gap: 12,
}

const panelTitleStyle: CSSProperties = {
  fontSize: 14,
  fontWeight: 700,
}

const panelSubtitleStyle: CSSProperties = {
  marginTop: 4,
  fontSize: 12,
  color: '#64748b',
  lineHeight: 1.45,
}

const closeButtonStyle: CSSProperties = {
  border: 'none',
  background: 'transparent',
  color: '#64748b',
  fontSize: 12,
  cursor: 'pointer',
}

const sectionStyle: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  gap: 10,
  marginTop: 14,
  minHeight: 0,
  flex: 1,
}

const sectionHeaderStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'flex-start',
  justifyContent: 'space-between',
  gap: 12,
}

const labelStyle: CSSProperties = {
  fontSize: 11,
  textTransform: 'uppercase',
  letterSpacing: '0.08em',
  color: '#64748b',
}

const providerRowStyle: CSSProperties = {
  display: 'flex',
  flexWrap: 'wrap',
  gap: 6,
}

const inputStyle: CSSProperties = {
  width: '100%',
  borderRadius: 8,
  border: '1px solid rgba(148, 163, 184, 0.22)',
  background: '#f8fafc',
  color: '#0f172a',
  padding: '9px 12px',
  fontSize: 13,
  boxSizing: 'border-box',
}

const textareaStyle: CSSProperties = {
  ...inputStyle,
  minHeight: 112,
  resize: 'vertical',
  fontFamily: 'inherit',
  lineHeight: 1.5,
}

const sheetPickerStyle: CSSProperties = {
  borderRadius: 10,
  border: '1px solid rgba(37, 99, 235, 0.18)',
  background: '#ffffff',
  boxShadow: '0 16px 40px rgba(15, 23, 42, 0.12)',
  padding: 8,
  display: 'flex',
  flexDirection: 'column',
  gap: 4,
}

const sheetPickerHeaderStyle: CSSProperties = {
  padding: '4px 6px 6px',
  fontSize: 11,
  fontWeight: 800,
  color: '#64748b',
}

const sheetOptionStyle: CSSProperties = {
  border: 'none',
  borderRadius: 8,
  background: '#f8fafc',
  color: '#0f172a',
  padding: '8px 10px',
  fontSize: 12,
  textAlign: 'left',
  cursor: 'pointer',
}

const sheetPickerEmptyStyle: CSSProperties = {
  padding: '8px 10px',
  fontSize: 12,
  color: '#94a3b8',
}

const rangeBindingBarStyle: CSSProperties = {
  borderRadius: 8,
  border: '1px solid rgba(37, 99, 235, 0.16)',
  background: '#eff6ff',
  color: '#1d4ed8',
  padding: '8px 10px',
  fontSize: 12,
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  gap: 10,
}

const inlineTextButtonStyle: CSSProperties = {
  border: 'none',
  background: 'transparent',
  color: '#1d4ed8',
  fontSize: 12,
  fontWeight: 800,
  cursor: 'pointer',
  padding: 0,
  flexShrink: 0,
}

const selectStyle: CSSProperties = {
  ...inputStyle,
  appearance: 'none',
}

const helperTextStyle: CSSProperties = {
  fontSize: 12,
  color: '#64748b',
  lineHeight: 1.45,
}

const actionsRowStyle: CSSProperties = {
  display: 'flex',
  gap: 8,
  marginTop: 6,
}

const commandRowStyle: CSSProperties = {
  display: 'flex',
  gap: 8,
}

const primaryButtonStyle: CSSProperties = {
  border: 'none',
  borderRadius: 8,
  background: '#2563eb',
  color: '#eff6ff',
  padding: '10px 14px',
  fontSize: 12,
  fontWeight: 600,
  cursor: 'pointer',
}

const secondaryButtonStyle: CSSProperties = {
  border: '1px solid rgba(148, 163, 184, 0.24)',
  borderRadius: 8,
  background: '#ffffff',
  color: '#334155',
  padding: '10px 14px',
  fontSize: 12,
  fontWeight: 600,
  cursor: 'pointer',
}

const disabledButtonStyle: CSSProperties = {
  ...secondaryButtonStyle,
  opacity: 0.48,
  cursor: 'not-allowed',
}

const dangerButtonStyle: CSSProperties = {
  border: '1px solid rgba(239, 68, 68, 0.25)',
  borderRadius: 8,
  background: '#fff5f5',
  color: '#dc2626',
  padding: '10px 14px',
  fontSize: 12,
  fontWeight: 600,
  cursor: 'pointer',
}

const tagButtonStyle: CSSProperties = {
  border: 'none',
  borderRadius: 999,
  padding: '7px 12px',
  fontSize: 12,
  cursor: 'pointer',
}

const statusCardStyle: CSSProperties = {
  borderRadius: 8,
  background: '#f8fafc',
  border: '1px solid rgba(148, 163, 184, 0.18)',
  padding: '12px 14px',
}

const statusTitleStyle: CSSProperties = {
  fontSize: 12,
  fontWeight: 700,
  color: '#0f172a',
}

const hintCardStyle: CSSProperties = {
  borderRadius: 8,
  background: '#eff6ff',
  color: '#1d4ed8',
  padding: '10px 12px',
  fontSize: 12,
  lineHeight: 1.5,
}

const historyAreaStyle: CSSProperties = {
  flex: 1,
  minHeight: 260,
  overflowY: 'auto',
  display: 'flex',
  flexDirection: 'column',
  gap: 12,
  paddingRight: 2,
}

const loadMoreButtonStyle: CSSProperties = {
  border: '1px dashed rgba(148, 163, 184, 0.34)',
  borderRadius: 12,
  background: '#ffffff',
  color: '#475569',
  padding: '10px 12px',
  fontSize: 12,
  fontWeight: 700,
  cursor: 'pointer',
}

const emptyHistoryStyle: CSSProperties = {
  borderRadius: 12,
  border: '1px dashed rgba(148, 163, 184, 0.32)',
  background: '#f8fafc',
  color: '#64748b',
  padding: '18px 16px',
  fontSize: 12,
  lineHeight: 1.6,
}

const executionCardStyle: CSSProperties = {
  borderRadius: 14,
  border: '1px solid rgba(37, 99, 235, 0.18)',
  background: '#ffffff',
  padding: 14,
  boxShadow: '0 10px 28px rgba(15, 23, 42, 0.06)',
  display: 'flex',
  flexDirection: 'column',
  gap: 12,
}

const executionHeaderStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'flex-start',
  justifyContent: 'space-between',
  gap: 12,
}

const executionTitleWrapStyle: CSSProperties = {
  minWidth: 0,
}

const executionCommandStyle: CSSProperties = {
  fontSize: 13,
  fontWeight: 800,
  color: '#0f172a',
  lineHeight: 1.45,
  wordBreak: 'break-word',
  userSelect: 'text',
}

const recordSummaryStyle: CSSProperties = {
  borderRadius: 10,
  border: '1px solid rgba(148, 163, 184, 0.16)',
  background: '#f8fafc',
  color: '#475569',
  padding: '9px 10px',
  display: 'flex',
  alignItems: 'flex-start',
  justifyContent: 'space-between',
  gap: 10,
  fontSize: 12,
  lineHeight: 1.45,
  userSelect: 'text',
}

const executionMetaStyle: CSSProperties = {
  marginTop: 4,
  fontSize: 11,
  color: '#94a3b8',
}

const executionStatusBadgeStyle: CSSProperties = {
  borderRadius: 999,
  padding: '6px 10px',
  fontSize: 12,
  fontWeight: 800,
  flexShrink: 0,
}

const metricGridStyle: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: 'repeat(7, minmax(0, 1fr))',
  gap: 8,
}

const metricTileStyle: CSSProperties = {
  borderRadius: 10,
  border: '1px solid rgba(148, 163, 184, 0.18)',
  background: '#f8fafc',
  padding: '9px 8px',
  display: 'flex',
  flexDirection: 'column',
  gap: 2,
  fontSize: 11,
  minWidth: 0,
}

const metricValueStyle: CSSProperties = {
  fontSize: 16,
  fontWeight: 900,
}

const operationTableStyle: CSSProperties = {
  borderRadius: 10,
  border: '1px solid rgba(148, 163, 184, 0.18)',
  overflow: 'hidden',
  background: '#ffffff',
}

const operationTableHeaderStyle: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: '54px 1.1fr 1.2fr 78px',
  gap: 8,
  padding: '9px 10px',
  background: '#f1f5f9',
  color: '#64748b',
  fontSize: 11,
  fontWeight: 800,
}

const operationRowStyle: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: '54px 1.1fr 1.2fr 78px',
  gap: 8,
  alignItems: 'center',
  padding: '9px 10px',
  borderTop: '1px solid rgba(148, 163, 184, 0.14)',
  fontSize: 12,
}

const operationBadgeStyle: CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  borderRadius: 999,
  background: '#eef2ff',
  color: '#3730a3',
  fontSize: 11,
  fontWeight: 800,
  padding: '4px 7px',
}

const operationNameStyle: CSSProperties = {
  color: '#0f172a',
  fontWeight: 700,
  minWidth: 0,
}

const operationScopeStyle: CSSProperties = {
  color: '#475569',
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap',
  minWidth: 0,
}

const operationStateStyle: CSSProperties = {
  fontSize: 12,
  fontWeight: 700,
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap',
}

const tableFooterStyle: CSSProperties = {
  padding: '9px 10px',
  borderTop: '1px solid rgba(148, 163, 184, 0.14)',
  color: '#64748b',
  fontSize: 12,
}

const verificationCardStyle: CSSProperties = {
  borderRadius: 10,
  border: '1px solid rgba(148, 163, 184, 0.18)',
  padding: '10px 12px',
}

const errorTextStyle: CSSProperties = {
  borderRadius: 10,
  background: '#fff5f5',
  color: '#b91c1c',
  padding: '10px 12px',
  fontSize: 12,
  lineHeight: 1.5,
  userSelect: 'text',
}

const recordActionsStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  gap: 10,
  paddingTop: 2,
}

const recordButtonGroupStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  flexShrink: 0,
}

const recordHintStyle: CSSProperties = {
  color: '#94a3b8',
  fontSize: 11,
  lineHeight: 1.4,
  textAlign: 'right',
}

const diffCardStyle: CSSProperties = {
  borderRadius: 12,
  border: '1px solid rgba(14, 165, 233, 0.22)',
  background: '#f0f9ff',
  padding: 12,
  display: 'flex',
  flexDirection: 'column',
  gap: 10,
}

const diffHeaderStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'flex-start',
  justifyContent: 'space-between',
  gap: 10,
}

const diffSubtitleStyle: CSSProperties = {
  marginTop: 4,
  fontSize: 11,
  color: '#0369a1',
}

const diffBadgeStyle: CSSProperties = {
  borderRadius: 999,
  background: '#e0f2fe',
  color: '#0369a1',
  padding: '6px 10px',
  fontSize: 12,
  fontWeight: 800,
  flexShrink: 0,
}

const diffSheetNoticeStyle: CSSProperties = {
  borderRadius: 8,
  background: '#ffffff',
  border: '1px solid rgba(14, 165, 233, 0.18)',
  color: '#075985',
  padding: '8px 10px',
  fontSize: 12,
  lineHeight: 1.45,
}

const diffEmptyStyle: CSSProperties = {
  borderRadius: 8,
  background: '#ffffff',
  color: '#64748b',
  padding: '9px 10px',
  fontSize: 12,
}

const diffTableStyle: CSSProperties = {
  borderRadius: 10,
  border: '1px solid rgba(14, 165, 233, 0.18)',
  overflow: 'hidden',
  background: '#ffffff',
}

const diffTableHeaderStyle: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: '1fr 70px 1.2fr 1.2fr',
  gap: 8,
  padding: '9px 10px',
  background: '#e0f2fe',
  color: '#0369a1',
  fontSize: 11,
  fontWeight: 900,
}

const diffRowStyle: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: '1fr 70px 1.2fr 1.2fr',
  gap: 8,
  alignItems: 'center',
  padding: '9px 10px',
  borderTop: '1px solid rgba(14, 165, 233, 0.14)',
  fontSize: 12,
}

const diffCellRefStyle: CSSProperties = {
  color: '#0f172a',
  fontWeight: 800,
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap',
}

const diffKindStyle: CSSProperties = {
  borderRadius: 999,
  background: '#f1f5f9',
  color: '#475569',
  padding: '4px 7px',
  fontSize: 11,
  fontWeight: 800,
  textAlign: 'center',
}

const diffValueStyle: CSSProperties = {
  color: '#334155',
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap',
  minWidth: 0,
}

const diffFooterStyle: CSSProperties = {
  color: '#64748b',
  fontSize: 11,
  lineHeight: 1.45,
}

const diffActionsStyle: CSSProperties = {
  display: 'flex',
  justifyContent: 'flex-end',
}

const compactButtonStyle: CSSProperties = {
  border: '1px solid rgba(14, 165, 233, 0.22)',
  borderRadius: 999,
  background: '#ffffff',
  color: '#0369a1',
  padding: '5px 9px',
  fontSize: 11,
  fontWeight: 800,
  cursor: 'pointer',
}

const diffModalStyle: CSSProperties = {
  width: 'min(1080px, calc(100vw - 48px))',
  maxHeight: 'min(88vh, 920px)',
  overflow: 'hidden',
  borderRadius: 16,
  border: '1px solid rgba(148, 163, 184, 0.22)',
  background: '#ffffff',
  boxShadow: '0 30px 90px rgba(15, 23, 42, 0.30)',
  padding: 20,
  color: '#0f172a',
  display: 'flex',
  flexDirection: 'column',
  gap: 14,
}

const diffOverviewGridStyle: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: 'repeat(4, minmax(0, 1fr))',
  gap: 10,
}

const diffOverviewTileStyle: CSSProperties = {
  borderRadius: 12,
  border: '1px solid rgba(14, 165, 233, 0.18)',
  background: '#f0f9ff',
  padding: '12px 14px',
  display: 'flex',
  flexDirection: 'column',
  gap: 4,
  color: '#0369a1',
  fontSize: 12,
  fontWeight: 700,
}

const diffToolbarStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  gap: 12,
}

const fullDiffTableStyle: CSSProperties = {
  borderRadius: 12,
  border: '1px solid rgba(148, 163, 184, 0.20)',
  overflow: 'auto',
  minHeight: 220,
  flex: 1,
  background: '#ffffff',
}

const fullDiffTableHeaderStyle: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: '180px 86px minmax(180px, 1fr) minmax(180px, 1fr)',
  gap: 10,
  padding: '10px 12px',
  background: '#f8fafc',
  color: '#64748b',
  fontSize: 11,
  fontWeight: 900,
  position: 'sticky',
  top: 0,
  zIndex: 1,
}

const fullDiffRowStyle: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: '180px 86px minmax(180px, 1fr) minmax(180px, 1fr)',
  gap: 10,
  alignItems: 'center',
  padding: '10px 12px',
  borderTop: '1px solid rgba(148, 163, 184, 0.14)',
  fontSize: 12,
}

const fullDiffValueStyle: CSSProperties = {
  color: '#334155',
  whiteSpace: 'pre-wrap',
  wordBreak: 'break-word',
  lineHeight: 1.45,
  minWidth: 0,
  userSelect: 'text',
}

const fullDiffEmptyStyle: CSSProperties = {
  padding: '18px 14px',
  color: '#64748b',
  fontSize: 12,
}

const planCardStyle: CSSProperties = {
  borderRadius: 8,
  border: '1px solid rgba(37, 99, 235, 0.24)',
  background: '#f8fbff',
  padding: 14,
  display: 'flex',
  flexDirection: 'column',
  gap: 12,
  overflowY: 'auto',
  maxHeight: 360,
}

const planCardHeaderStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'flex-start',
  justifyContent: 'space-between',
  gap: 12,
}

const planTitleStyle: CSSProperties = {
  fontSize: 13,
  fontWeight: 700,
  color: '#1d4ed8',
}

const planStateBadgeStyle: CSSProperties = {
  borderRadius: 999,
  background: '#dbeafe',
  color: '#1d4ed8',
  padding: '6px 10px',
  fontSize: 12,
  fontWeight: 700,
  flexShrink: 0,
}

const planMetaStyle: CSSProperties = {
  marginTop: 4,
  fontSize: 11,
  color: '#475569',
  wordBreak: 'break-all',
}

const planBodyStyle: CSSProperties = {
  fontSize: 12,
  color: '#1e3a8a',
  lineHeight: 1.6,
  whiteSpace: 'pre-wrap',
  userSelect: 'text',
}

const statGridStyle: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: 'repeat(5, minmax(0, 1fr))',
  gap: 6,
}

const statPillStyle: CSSProperties = {
  borderRadius: 8,
  border: '1px solid rgba(37, 99, 235, 0.18)',
  background: '#ffffff',
  padding: '8px 8px',
  fontSize: 11,
  display: 'flex',
  flexDirection: 'column',
  gap: 2,
  minWidth: 0,
}

const statValueStyle: CSSProperties = {
  fontSize: 15,
  fontWeight: 800,
}

const detailBlockStyle: CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  gap: 6,
}

const detailTitleStyle: CSSProperties = {
  fontSize: 12,
  fontWeight: 700,
  color: '#0f172a',
}

const detailListStyle: CSSProperties = {
  display: 'grid',
  gap: 6,
}

const detailItemStyle: CSSProperties = {
  borderRadius: 8,
  border: '1px solid rgba(148, 163, 184, 0.18)',
  padding: '8px 10px',
  fontSize: 12,
  lineHeight: 1.45,
  wordBreak: 'break-word',
  userSelect: 'text',
}

const summaryTitleStyle: CSSProperties = {
  fontSize: 12,
  fontWeight: 700,
  color: '#0f172a',
}

const summaryTextStyle: CSSProperties = {
  marginTop: 6,
  fontSize: 12,
  color: '#475569',
  lineHeight: 1.55,
  whiteSpace: 'pre-wrap',
  wordBreak: 'break-word',
  userSelect: 'text',
}

const chipGridStyle: CSSProperties = {
  display: 'flex',
  flexWrap: 'wrap',
  gap: 8,
}

const chipButtonStyle: CSSProperties = {
  borderRadius: 999,
  border: '1px solid rgba(148, 163, 184, 0.22)',
  padding: '8px 12px',
  fontSize: 12,
  cursor: 'pointer',
}

const emptyStateStyle: CSSProperties = {
  borderRadius: 8,
  background: '#f8fafc',
  color: '#64748b',
  padding: '12px 14px',
  fontSize: 12,
  lineHeight: 1.5,
}

const savedModelListStyle: CSSProperties = {
  display: 'grid',
  gap: 10,
}

const savedModelCardStyle: CSSProperties = {
  borderRadius: 8,
  border: '1px solid rgba(148, 163, 184, 0.2)',
  background: '#ffffff',
  padding: 14,
}

const savedModelHeaderStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'flex-start',
  justifyContent: 'space-between',
  gap: 10,
}

const savedModelTitleStyle: CSSProperties = {
  fontSize: 14,
  fontWeight: 700,
  color: '#0f172a',
}

const savedModelMetaStyle: CSSProperties = {
  marginTop: 4,
  fontSize: 12,
  color: '#64748b',
  lineHeight: 1.45,
  wordBreak: 'break-all',
}

const activeBadgeStyle: CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  minWidth: 42,
  height: 24,
  borderRadius: 999,
  background: '#dbeafe',
  color: '#1d4ed8',
  fontSize: 12,
  fontWeight: 700,
}

const savedModelActionsStyle: CSSProperties = {
  display: 'flex',
  gap: 8,
  marginTop: 12,
}

const dividerStyle: CSSProperties = {
  marginTop: 16,
  borderTop: '1px solid rgba(148, 163, 184, 0.16)',
}
