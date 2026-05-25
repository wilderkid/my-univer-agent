export type CellValue = string | number | boolean | null
export type RowConditions = Record<string, unknown>
export type StructuredRecord = Record<string, unknown>

export interface RangeBounds {
  startRow: number
  startColumn: number
  endRow: number
  endColumn: number
}

export interface RangeLike {
  getA1Notation(withSheet?: boolean): string
  getValues(): unknown[][]
  getFormulas(): string[][]
  getCellDataGrid(): Array<Array<Record<string, unknown> | null>>
  getRow(): number
  getColumn(): number
  getHeight(): number
  getWidth(): number
  getRange(): RangeBounds
  setFormula(value: string): void
  setValue(value: string | number | boolean): void
  setValues(values: Array<Array<string | number | boolean | Record<string, unknown> | null>>): void
  clear(options?: { contentsOnly?: boolean; formatOnly?: boolean }): void
  setFontWeight(value: string): void
  setFontStyle(value: string): void
  setFontSize(value: number): void
  setFontColor(value: string): void
  setFontFamily?(value: string): void
  setBackground(value: string): void
  setFontLine(value: string): void
  setHorizontalAlignment(value: string): void
  setVerticalAlignment?(value: string): void
  setTextRotation?(value: number): void
  setBorder(type: unknown, style: unknown, color?: string): void
  setWrapStrategy(strategy: unknown): void
  breakApart(): void
  merge(options?: unknown): void
}

export interface SheetLike {
  getSheetId(): string
  getSheetName(): string
  getLastRow(): number
  getLastColumn(): number
  getRange(a1: string): RangeLike
  getRange(row: number, column: number): RangeLike
  getRange(row: number, column: number, numRows: number): RangeLike
  getRange(row: number, column: number, numRows: number, numColumns: number): RangeLike
  getSelection(): { getActiveRangeList(): RangeLike[] } | null
  insertRows(rowIndex: number, count: number): void
  deleteRows(rowIndex: number, count: number): void
  hideRows?(rowIndex: number, count?: number): void
  showRows?(rowIndex: number, count?: number): void
  unhideRow?(row: RangeLike): void
  insertColumns(columnIndex: number, count: number): void
  deleteColumns(columnIndex: number, count: number): void
  hideColumns?(columnIndex: number, count?: number): void
  showColumns?(columnIndex: number, count?: number): void
  unhideColumn?(column: RangeLike): void
  getColumnWidth(columnIndex: number): number
  setColumnWidth(columnIndex: number, width: number): void
  setColumnWidths?(startColumn: number, count: number, width: number): void
  getRowHeight(rowIndex: number): number
  setRowHeight(rowIndex: number, height: number): void
  setRowHeights?(startRow: number, count: number, height: number): void
  setRowHeightsForced?(startRow: number, count: number, height: number): void
  setRowAutoHeight?(startRow: number, count: number): void
  setRangesAutoHeight?(ranges: RangeBounds[]): void
  setFrozenRows?(rows: number): void
  setFrozenColumns?(columns: number): void
  setHiddenGridlines?(hidden: boolean): void
  setTabColor?(color: string): void
  getMergedRanges(): RangeLike[]
  setName(name: string): void
}

export interface WorkbookLike {
  getSheets(): SheetLike[]
  getActiveSheet(): SheetLike | null
  getActiveCell(): { getA1Notation(): string } | null
  getActiveRange(): RangeLike | null
  getSheetByName(name: string): SheetLike | null
  insertSheet(name: string): SheetLike
  deleteSheet(sheet: SheetLike | string): boolean
}

export interface TableColumnDescriptor {
  header: string
  normalizedHeader: string
  columnIndex: number
  relativeColumnIndex: number
  columnLetter: string
  headerCell: string
  sampleValues: CellValue[]
}

export interface TableDescriptor {
  sheetName: string
  range: string
  headerRowIndex: number
  headerRowNumber: number
  headerRowRange: string
  dataRange: string | null
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
  rowCount: number
  columnCount: number
  columns: TableColumnDescriptor[]
  previewRows: StructuredRecord[]
  primaryKeyCandidates: string[]
  dataRows: CellValue[][]
  dataFormulaRows: string[][]
}

export interface SheetSchema {
  sheetName: string
  usedRange: string | null
  tables: TableDescriptor[]
}

export interface WorkbookSchemaSnapshot {
  activeSheetName: string | null
  activeCell: string | null
  activeRange: string | null
  activeRangePosition: { row: number; column: number } | null
  selectedRanges: string[]
  sheets: SheetSchema[]
  tables: TableDescriptor[]
}

export interface ResolvedCondition {
  requestedHeader: string
  matchedHeader: string
  columnIndex: number
  expectedValue: unknown
  score: number
}

export interface RowMatch {
  rowIndex: number
  rowNumber: number
  record: StructuredRecord
  score: number
}

export interface TableHeaderMapping {
  sourceHeader: string
  sourceRelativeColumnIndex: number
  targetHeader: string
  targetRelativeColumnIndex: number
  score: number
}

export interface MatchHeaderAnalysis {
  sourceHeader: string
  targetHeader: string
  score: number
  sharedValueCount: number
  sourceDistinctCount: number
  targetDistinctCount: number
  sharedValueSamples: string[]
}

export interface DerivedCellComputation {
  targetHeader: string
  targetColumnLetter: string
  value: CellValue | string
  reason: string
}

export interface CellLocationResult {
  sheet_name: string
  table_range: string
  target_cell: string
  target_row_number: number
  target_column_header: string
  target_column_cell: string
  matched_row: StructuredRecord
  current_value: CellValue
  current_formula: string | null
  confidence: 'high' | 'medium'
}

export const TABLE_PREVIEW_ROWS = 3
export const CONTEXT_PREVIEW_MAX_ROWS = 12
export const CONTEXT_PREVIEW_MAX_COLUMNS = 8
export const MIN_HEADER_MATCH_SCORE = 55

export function colIndexToLetter(index: number): string {
  let letter = ''
  let n = index
  while (n >= 0) {
    letter = String.fromCharCode(65 + (n % 26)) + letter
    n = Math.floor(n / 26) - 1
  }
  return letter
}

export function tableColumnLetterToIndex(columnLetter: string): number {
  let index = 0
  const normalized = columnLetter.trim().toUpperCase()
  for (let i = 0; i < normalized.length; i++) {
    index = index * 26 + (normalized.charCodeAt(i) - 64)
  }
  return index - 1
}

export function rangeToA1(startRow: number, startColumn: number, endRow: number, endColumn: number): string {
  return `${colIndexToLetter(startColumn)}${startRow + 1}:${colIndexToLetter(endColumn)}${endRow + 1}`
}

export function isFormulaString(value: unknown): value is string {
  return typeof value === 'string' && value.startsWith('=')
}

export function toCellValue(value: unknown): CellValue {
  if (value === null || value === undefined || value === '') return null
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value
  return String(value)
}

export function normalizeText(value: unknown): string {
  if (value === null || value === undefined) return ''
  return String(value)
    .trim()
    .toLowerCase()
    .replace(/[\u200b-\u200d\ufeff]/g, '')
    .replace(/[()[\]{}"'`~!@#$%^&*+=|\\/:;,.<>?_\-，。；：、！￥（）【】《》“”‘’\s]+/g, '')
}

export function normalizeSheetNameForMatch(value: unknown): string {
  if (value === null || value === undefined) return ''
  return String(value)
    .normalize('NFKC')
    .replace(/[\u00A0\u1680\u2000-\u200D\u202F\u205F\u3000\uFEFF]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
}

export function resolveSnapshotSheetName(
  snapshot: WorkbookSchemaSnapshot,
  requestedSheetName: unknown,
): string | null {
  if (typeof requestedSheetName !== 'string' || !requestedSheetName.trim()) return null

  const exactMatch = snapshot.sheets.find((sheet) => sheet.sheetName === requestedSheetName.trim())
  if (exactMatch) return exactMatch.sheetName

  const normalizedRequested = normalizeSheetNameForMatch(requestedSheetName)
  if (!normalizedRequested) return null

  const normalizedMatches = snapshot.sheets.filter((sheet) =>
    normalizeSheetNameForMatch(sheet.sheetName) === normalizedRequested,
  )
  if (normalizedMatches.length === 1) return normalizedMatches[0].sheetName

  return null
}

export function findSheetByName(workbook: WorkbookLike, requestedSheetName: unknown): SheetLike | null {
  if (typeof requestedSheetName !== 'string' || !requestedSheetName.trim()) return null

  const exactMatch = workbook.getSheetByName(requestedSheetName.trim())
  if (exactMatch) return exactMatch

  const normalizedRequested = normalizeSheetNameForMatch(requestedSheetName)
  if (!normalizedRequested) return null

  const normalizedMatches = workbook.getSheets().filter((sheet) =>
    normalizeSheetNameForMatch(sheet.getSheetName()) === normalizedRequested,
  )
  if (normalizedMatches.length === 1) return normalizedMatches[0]

  return null
}

function longestCommonSubstringLength(left: string, right: string): number {
  if (!left || !right) return 0
  const matrix = Array.from({ length: left.length + 1 }, () => Array<number>(right.length + 1).fill(0))
  let longest = 0

  for (let i = 1; i <= left.length; i++) {
    for (let j = 1; j <= right.length; j++) {
      if (left[i - 1] === right[j - 1]) {
        matrix[i][j] = matrix[i - 1][j - 1] + 1
        if (matrix[i][j] > longest) {
          longest = matrix[i][j]
        }
      }
    }
  }

  return longest
}

export function getHeaderMatchScore(candidate: string, query: string): number {
  const normalizedCandidate = normalizeText(candidate)
  const normalizedQuery = normalizeText(query)
  if (!normalizedCandidate || !normalizedQuery) return 0
  if (normalizedCandidate === normalizedQuery) return 100
  if (normalizedCandidate.includes(normalizedQuery) || normalizedQuery.includes(normalizedCandidate)) {
    return 82 - Math.min(20, Math.abs(normalizedCandidate.length - normalizedQuery.length))
  }

  const commonLength = longestCommonSubstringLength(normalizedCandidate, normalizedQuery)
  const ratio = commonLength / Math.max(normalizedCandidate.length, normalizedQuery.length)
  return ratio >= 0.6 ? Math.round(ratio * 70) : 0
}

export function getValueMatchScore(actualValue: unknown, expectedValue: unknown): number {
  if (expectedValue === null || expectedValue === undefined || expectedValue === '') {
    return actualValue === null || actualValue === undefined || actualValue === '' ? 100 : 0
  }

  const actualText = normalizeText(actualValue)
  const expectedText = normalizeText(expectedValue)
  if (!actualText || !expectedText) return 0
  if (actualText === expectedText) return 100
  if (actualText.includes(expectedText) || expectedText.includes(actualText)) return 72
  return 0
}

export function parseNumericValue(value: unknown): number | null {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null
  }

  if (typeof value !== 'string') {
    return null
  }

  const trimmed = value.trim()
  if (!trimmed) return null

  const negative = trimmed.startsWith('(') && trimmed.endsWith(')')
  const normalized = trimmed
    .replace(/[,$￥¥€£\s]/g, '')
    .replace(/[()]/g, '')
  const parsed = Number(normalized)
  if (!Number.isFinite(parsed)) return null

  return negative ? -parsed : parsed
}

export function getStrictValueMatchScore(actualValue: unknown, expectedValue: unknown): number {
  if (expectedValue === null || expectedValue === undefined || expectedValue === '') {
    return 0
  }

  const actualText = normalizeText(actualValue)
  const expectedText = normalizeText(expectedValue)
  if (!actualText || !expectedText) return 0
  return actualText === expectedText ? 100 : 0
}

function isNonEmptyCell(value: unknown, formula: string | undefined): boolean {
  return !!normalizeText(formula) || (value !== null && value !== undefined && value !== '')
}

function isRowBlank(values: unknown[][], formulas: string[][], rowIndex: number, columnCount: number): boolean {
  for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
    if (isNonEmptyCell(values[rowIndex]?.[columnIndex], formulas[rowIndex]?.[columnIndex])) {
      return false
    }
  }
  return true
}

function truncatePreviewRows(rows: StructuredRecord[], limit: number): StructuredRecord[] {
  return rows.slice(0, limit)
}

function getSheetValues(sheet: SheetLike): { values: unknown[][]; formulas: string[][]; lastRow: number; lastColumn: number } | null {
  const lastRow = sheet.getLastRow()
  const lastColumn = sheet.getLastColumn()
  if (lastRow < 0 || lastColumn < 0) return null

  const range = sheet.getRange(0, 0, lastRow + 1, lastColumn + 1)
  return {
    values: range.getValues(),
    formulas: range.getFormulas(),
    lastRow,
    lastColumn,
  }
}

function getNonEmptyColumnIndexes(rowValues: unknown[], rowFormulas: string[], columnCount: number): number[] {
  const indexes: number[] = []
  for (let columnIndex = 0; columnIndex < columnCount; columnIndex++) {
    if (isNonEmptyCell(rowValues[columnIndex], rowFormulas[columnIndex])) {
      indexes.push(columnIndex)
    }
  }
  return indexes
}

export function rowArrayToRecord(columns: TableColumnDescriptor[], rowValues: CellValue[]): StructuredRecord {
  const record: StructuredRecord = {}
  columns.forEach((column) => {
    record[column.header] = rowValues[column.relativeColumnIndex]
  })
  return record
}

function buildTableColumns(
  values: unknown[][],
  headerRowIndex: number,
  headerColumns: number[],
  segmentEnd: number,
): TableColumnDescriptor[] {
  const usedHeaders = new Map<string, number>()

  return headerColumns.map((columnIndex, relativeColumnIndex) => {
    const rawHeader = String(values[headerRowIndex]?.[columnIndex] ?? '').trim() || `Column ${colIndexToLetter(columnIndex)}`
    const duplicateCount = usedHeaders.get(rawHeader) ?? 0
    usedHeaders.set(rawHeader, duplicateCount + 1)
    const header = duplicateCount === 0 ? rawHeader : `${rawHeader} (${colIndexToLetter(columnIndex)})`
    const sampleValues: CellValue[] = []

    for (let rowIndex = headerRowIndex + 1; rowIndex <= segmentEnd && sampleValues.length < TABLE_PREVIEW_ROWS; rowIndex++) {
      const cellValue = toCellValue(values[rowIndex]?.[columnIndex])
      if (cellValue !== null) {
        sampleValues.push(cellValue)
      }
    }

    return {
      header,
      normalizedHeader: normalizeText(header),
      columnIndex,
      relativeColumnIndex,
      columnLetter: colIndexToLetter(columnIndex),
      headerCell: `${colIndexToLetter(columnIndex)}${headerRowIndex + 1}`,
      sampleValues,
    }
  })
}

function findPrimaryKeyCandidates(columns: TableColumnDescriptor[], dataRows: CellValue[][]): string[] {
  const candidates: string[] = []

  columns.forEach((column) => {
    const valuesInColumn = dataRows
      .map((row) => row[column.relativeColumnIndex])
      .filter((value): value is string | number | boolean => value !== null)
    if (valuesInColumn.length === 0) return

    const uniqueValues = new Set(valuesInColumn.map((value) => normalizeText(value)))
    if (uniqueValues.size === valuesInColumn.length) {
      candidates.push(column.header)
    }
  })

  return candidates.slice(0, 3)
}

function buildTableFromSegment(
  sheet: SheetLike,
  values: unknown[][],
  formulas: string[][],
  segmentStart: number,
  segmentEnd: number,
  lastColumn: number,
): TableDescriptor | null {
  const scanEnd = Math.min(segmentEnd, segmentStart + 4)
  let headerRowIndex = segmentStart
  let bestHeaderScore = -1
  let bestHeaderCellCount = 0

  for (let candidateRow = segmentStart; candidateRow <= scanEnd; candidateRow++) {
    const usedColumns = getNonEmptyColumnIndexes(values[candidateRow] ?? [], formulas[candidateRow] ?? [], lastColumn + 1)
    if (usedColumns.length === 0) continue

    const textCellCount = usedColumns.filter((columnIndex) => typeof values[candidateRow]?.[columnIndex] === 'string').length
    const headerScore = (usedColumns.length >= 2 ? 10 : 0) + textCellCount * 2
    if (headerScore > bestHeaderScore || (headerScore === bestHeaderScore && usedColumns.length > bestHeaderCellCount)) {
      headerRowIndex = candidateRow
      bestHeaderScore = headerScore
      bestHeaderCellCount = usedColumns.length
    }
  }

  const headerColumns = getNonEmptyColumnIndexes(values[headerRowIndex] ?? [], formulas[headerRowIndex] ?? [], lastColumn + 1)
  if (headerColumns.length === 0) return null

  const startColumn = headerColumns[0]
  const endColumn = headerColumns[headerColumns.length - 1]
  const columns = buildTableColumns(values, headerRowIndex, headerColumns, segmentEnd)
  const dataStartRow = headerRowIndex + 1
  const dataRows: CellValue[][] = []
  const dataFormulaRows: string[][] = []
  const previewRows: StructuredRecord[] = []

  for (let rowIndex = dataStartRow; rowIndex <= segmentEnd; rowIndex++) {
    const rowValues = columns.map((column) => toCellValue(values[rowIndex]?.[column.columnIndex]))
    const rowFormulas = columns.map((column) => formulas[rowIndex]?.[column.columnIndex] ?? '')
    dataRows.push(rowValues)
    dataFormulaRows.push(rowFormulas)
    previewRows.push(rowArrayToRecord(columns, rowValues))
  }

  return {
    sheetName: sheet.getSheetName(),
    range: rangeToA1(headerRowIndex, startColumn, segmentEnd, endColumn),
    headerRowIndex,
    headerRowNumber: headerRowIndex + 1,
    headerRowRange: rangeToA1(headerRowIndex, startColumn, headerRowIndex, endColumn),
    dataRange: dataStartRow <= segmentEnd ? rangeToA1(dataStartRow, startColumn, segmentEnd, endColumn) : null,
    startRow: headerRowIndex,
    endRow: segmentEnd,
    startColumn,
    endColumn,
    rowCount: Math.max(0, segmentEnd - headerRowIndex),
    columnCount: columns.length,
    columns,
    previewRows: truncatePreviewRows(previewRows, TABLE_PREVIEW_ROWS),
    primaryKeyCandidates: findPrimaryKeyCandidates(columns, dataRows),
    dataRows,
    dataFormulaRows,
  }
}

export function detectTablesInSheet(sheet: SheetLike): TableDescriptor[] {
  const sheetData = getSheetValues(sheet)
  if (!sheetData) return []

  const { values, formulas, lastRow, lastColumn } = sheetData
  const tables: TableDescriptor[] = []
  let rowIndex = 0

  while (rowIndex <= lastRow) {
    if (isRowBlank(values, formulas, rowIndex, lastColumn + 1)) {
      rowIndex++
      continue
    }

    const segmentStart = rowIndex
    while (rowIndex <= lastRow && !isRowBlank(values, formulas, rowIndex, lastColumn + 1)) {
      rowIndex++
    }

    const segmentEnd = rowIndex - 1
    const table = buildTableFromSegment(sheet, values, formulas, segmentStart, segmentEnd, lastColumn)
    if (table) {
      tables.push(table)
    }
  }

  return tables
}

export function extractWorkbookSchema(workbook: WorkbookLike): WorkbookSchemaSnapshot {
  const sheets = workbook.getSheets()
  const activeSheet = workbook.getActiveSheet()
  const activeCell = workbook.getActiveCell()
  const activeRange = workbook.getActiveRange()
  const selectedRanges = activeSheet?.getSelection()?.getActiveRangeList() ?? []

  const sheetSchemas = sheets.map((sheet) => {
    const lastRow = sheet.getLastRow()
    const lastColumn = sheet.getLastColumn()
    const usedRange = lastRow >= 0 && lastColumn >= 0
      ? `A1:${colIndexToLetter(lastColumn)}${lastRow + 1}`
      : null
    return {
      sheetName: sheet.getSheetName(),
      usedRange,
      tables: detectTablesInSheet(sheet),
    } satisfies SheetSchema
  })

  return {
    activeSheetName: activeSheet?.getSheetName() ?? null,
    activeCell: activeCell?.getA1Notation() ?? null,
    activeRange: activeRange?.getA1Notation() ?? null,
    activeRangePosition: activeRange
      ? { row: activeRange.getRow(), column: activeRange.getColumn() }
      : null,
    selectedRanges: selectedRanges.map((range) => range.getA1Notation()),
    sheets: sheetSchemas,
    tables: sheetSchemas.flatMap((sheet) => sheet.tables),
  }
}

export function summarizeTable(table: TableDescriptor): Record<string, unknown> {
  return {
    sheet_name: table.sheetName,
    range: table.range,
    header_row_number: table.headerRowNumber,
    header_row_range: table.headerRowRange,
    data_range: table.dataRange,
    row_count: table.rowCount,
    column_count: table.columnCount,
    columns: table.columns.map((column) => ({
      header: column.header,
      column_letter: column.columnLetter,
      header_cell: column.headerCell,
      sample_values: column.sampleValues,
    })),
    primary_key_candidates: table.primaryKeyCandidates,
    preview_rows: table.previewRows,
  }
}

export function resolveTable(
  snapshot: WorkbookSchemaSnapshot,
  args: Record<string, unknown>,
): TableDescriptor | null {
  const explicitSheetName = resolveSnapshotSheetName(snapshot, args.sheet_name) ?? undefined
  const explicitTableRange = typeof args.table_range === 'string' && args.table_range.trim()
    ? args.table_range.trim().toUpperCase()
    : undefined

  const tables = explicitSheetName
    ? snapshot.tables.filter((table) => table.sheetName === explicitSheetName)
    : snapshot.tables

  if (explicitTableRange) {
    return tables.find((table) => table.range.toUpperCase() === explicitTableRange) ?? null
  }

  if (explicitSheetName) {
    return tables[0] ?? null
  }

  const activeSheetName = snapshot.activeSheetName
  const activePosition = snapshot.activeRangePosition
  const activeSheetTables = activeSheetName
    ? tables.filter((table) => table.sheetName === activeSheetName)
    : tables

  if (activePosition) {
    const activeTable = activeSheetTables.find((table) =>
      activePosition.row >= table.startRow
      && activePosition.row <= table.endRow
      && activePosition.column >= table.startColumn
      && activePosition.column <= table.endColumn,
    )
    if (activeTable) {
      return activeTable
    }
  }

  return activeSheetTables[0] ?? tables[0] ?? null
}

export function resolveTableByPrefix(
  snapshot: WorkbookSchemaSnapshot,
  args: Record<string, unknown>,
  prefix: 'source' | 'target',
): TableDescriptor | null {
  const prefixedArgs: Record<string, unknown> = {}
  const sheetNameKey = `${prefix}_sheet_name`
  const tableRangeKey = `${prefix}_table_range`

  if (args[sheetNameKey] !== undefined) {
    prefixedArgs.sheet_name = args[sheetNameKey]
  }

  if (args[tableRangeKey] !== undefined) {
    prefixedArgs.table_range = args[tableRangeKey]
  }

  return resolveTable(snapshot, prefixedArgs)
}

export function getColumnMatches(table: TableDescriptor, query: string): Array<{ column: TableColumnDescriptor; score: number }> {
  return table.columns
    .map((column) => ({
      column,
      score: getHeaderMatchScore(column.header, query),
    }))
    .filter((match) => match.score >= MIN_HEADER_MATCH_SCORE)
    .sort((left, right) => right.score - left.score)
}

export function getBestColumnMatch(table: TableDescriptor, query: string): { column: TableColumnDescriptor; score: number } | null {
  return getColumnMatches(table, query)[0] ?? null
}

export function resolveConditions(table: TableDescriptor, rowConditions: RowConditions): {
  resolved: ResolvedCondition[]
  unresolved: string[]
} {
  const resolved: ResolvedCondition[] = []
  const unresolved: string[] = []

  Object.entries(rowConditions).forEach(([requestedHeader, expectedValue]) => {
    const bestMatch = getBestColumnMatch(table, requestedHeader)
    if (!bestMatch) {
      unresolved.push(requestedHeader)
      return
    }

    resolved.push({
      requestedHeader,
      matchedHeader: bestMatch.column.header,
      columnIndex: bestMatch.column.columnIndex,
      expectedValue,
      score: bestMatch.score,
    })
  })

  return { resolved, unresolved }
}

export function findRowsByConditions(
  table: TableDescriptor,
  rowConditions: RowConditions,
  limit = 5,
): {
  matches: RowMatch[]
  resolved_conditions: ResolvedCondition[]
  unresolved_conditions: string[]
} {
  const { resolved, unresolved } = resolveConditions(table, rowConditions)
  if (resolved.length === 0) {
    return { matches: [], resolved_conditions: resolved, unresolved_conditions: unresolved }
  }

  const matches: RowMatch[] = []
  table.dataRows.forEach((rowValues, dataRowIndex) => {
    let totalScore = 0
    for (const condition of resolved) {
      const column = table.columns.find((entry) => entry.columnIndex === condition.columnIndex)
      if (!column) return
      const actualValue = rowValues[column.relativeColumnIndex]
      const score = getValueMatchScore(actualValue, condition.expectedValue)
      if (score === 0) {
        return
      }
      totalScore += score
    }

    const rowIndex = table.headerRowIndex + 1 + dataRowIndex
    matches.push({
      rowIndex,
      rowNumber: rowIndex + 1,
      record: rowArrayToRecord(table.columns, rowValues),
      score: Math.round(totalScore / resolved.length),
    })
  })

  return {
    matches: matches.sort((left, right) => right.score - left.score).slice(0, limit),
    resolved_conditions: resolved,
    unresolved_conditions: unresolved,
  }
}

export function locateTargetCell(
  snapshot: WorkbookSchemaSnapshot,
  args: Record<string, unknown>,
): {
  location: CellLocationResult | null
  resolved_conditions?: ResolvedCondition[]
  unresolved_conditions?: string[]
  ambiguity?: RowMatch[]
  error?: string
} {
  const table = resolveTable(snapshot, args)
  if (!table) {
    return { location: null, error: 'No semantic table was detected for the requested sheet or range.' }
  }

  const columnHeader = typeof args.column_header === 'string' ? args.column_header.trim() : ''
  const rowConditions = readRowConditions(args.row_conditions)
  if (!columnHeader) {
    return { location: null, error: 'column_header is required.' }
  }

  const bestColumn = getBestColumnMatch(table, columnHeader)
  if (!bestColumn) {
    return {
      location: null,
      error: `No matching target column was found for "${columnHeader}" in table ${table.range}.`,
    }
  }

  const rowSearch = findRowsByConditions(table, rowConditions)
  if (rowSearch.unresolved_conditions.length > 0) {
    return {
      location: null,
      resolved_conditions: rowSearch.resolved_conditions,
      unresolved_conditions: rowSearch.unresolved_conditions,
      error: `Some row condition headers could not be matched: ${rowSearch.unresolved_conditions.join(', ')}`,
    }
  }

  if (rowSearch.matches.length === 0) {
    return {
      location: null,
      resolved_conditions: rowSearch.resolved_conditions,
      unresolved_conditions: rowSearch.unresolved_conditions,
      error: 'No row matched the provided row_conditions.',
    }
  }

  const bestRow = rowSearch.matches[0]
  const secondRow = rowSearch.matches[1]
  if (secondRow && bestRow.score === secondRow.score) {
    return {
      location: null,
      resolved_conditions: rowSearch.resolved_conditions,
      unresolved_conditions: rowSearch.unresolved_conditions,
      ambiguity: rowSearch.matches,
      error: 'The target row is ambiguous. Multiple rows match the row_conditions with the same confidence.',
    }
  }

  const dataRowOffset = bestRow.rowIndex - table.headerRowIndex - 1
  const currentValue = table.dataRows[dataRowOffset]?.[bestColumn.column.relativeColumnIndex] ?? null
  const currentFormula = table.dataFormulaRows[dataRowOffset]?.[bestColumn.column.relativeColumnIndex] || null

  return {
    location: {
      sheet_name: table.sheetName,
      table_range: table.range,
      target_cell: `${bestColumn.column.columnLetter}${bestRow.rowNumber}`,
      target_row_number: bestRow.rowNumber,
      target_column_header: bestColumn.column.header,
      target_column_cell: bestColumn.column.headerCell,
      matched_row: bestRow.record,
      current_value: currentValue,
      current_formula: currentFormula,
      confidence: bestColumn.score >= 90 && bestRow.score >= 90 ? 'high' : 'medium',
    },
    resolved_conditions: rowSearch.resolved_conditions,
    unresolved_conditions: rowSearch.unresolved_conditions,
  }
}

export function readRowConditions(value: unknown): RowConditions {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return {}
  }
  return value as RowConditions
}

export function readRecords(value: unknown): StructuredRecord[] {
  if (!Array.isArray(value)) return []
  return value.filter((entry): entry is StructuredRecord => !!entry && typeof entry === 'object' && !Array.isArray(entry))
}

export function buildHeaderMappings(sourceTable: TableDescriptor, targetTable: TableDescriptor): TableHeaderMapping[] {
  return sourceTable.columns
    .map((sourceColumn) => {
      const bestMatch = getBestColumnMatch(targetTable, sourceColumn.header)
      if (!bestMatch) return null

      return {
        sourceHeader: sourceColumn.header,
        sourceRelativeColumnIndex: sourceColumn.relativeColumnIndex,
        targetHeader: bestMatch.column.header,
        targetRelativeColumnIndex: bestMatch.column.relativeColumnIndex,
        score: bestMatch.score,
      } satisfies TableHeaderMapping
    })
    .filter((mapping): mapping is TableHeaderMapping => !!mapping)
    .sort((left, right) => left.targetRelativeColumnIndex - right.targetRelativeColumnIndex)
}

function looksLikeKeyHeader(header: string): boolean {
  const normalized = normalizeText(header)
  return ['id', 'code', 'name', 'project', 'item', 'sku', 'month', 'date', '编号', '编码', '名称', '项目', '产品', '日期', '月份']
    .some((keyword) => normalized.includes(normalizeText(keyword)))
}

export function resolveSyncMatchHeaders(
  sourceTable: TableDescriptor,
  targetTable: TableDescriptor,
  headerMappings: TableHeaderMapping[],
  requestedHeaders: string[],
): { pairs: TableHeaderMapping[]; unresolved: string[]; autoSelected: boolean } {
  if (requestedHeaders.length > 0) {
    const pairs: TableHeaderMapping[] = []
    const unresolved: string[] = []

    requestedHeaders.forEach((header) => {
      const sourceMatch = getBestColumnMatch(sourceTable, header)
      const targetMatch = getBestColumnMatch(targetTable, header)
      if (!sourceMatch || !targetMatch) {
        unresolved.push(header)
        return
      }

      pairs.push({
        sourceHeader: sourceMatch.column.header,
        sourceRelativeColumnIndex: sourceMatch.column.relativeColumnIndex,
        targetHeader: targetMatch.column.header,
        targetRelativeColumnIndex: targetMatch.column.relativeColumnIndex,
        score: Math.min(sourceMatch.score, targetMatch.score),
      })
    })

    return { pairs, unresolved, autoSelected: false }
  }

  const keyMappedPairs = headerMappings.filter((mapping) =>
    sourceTable.primaryKeyCandidates.includes(mapping.sourceHeader)
    || targetTable.primaryKeyCandidates.includes(mapping.targetHeader)
    || looksLikeKeyHeader(mapping.sourceHeader)
    || looksLikeKeyHeader(mapping.targetHeader),
  )

  const analyzedKeyPairs = analyzeMatchHeaderPairs(sourceTable, targetTable, keyMappedPairs)
    .filter((entry) => entry.sharedValueCount > 0 && entry.score >= 70)
    .sort((left, right) => {
      if (right.sharedValueCount !== left.sharedValueCount) {
        return right.sharedValueCount - left.sharedValueCount
      }
      return right.score - left.score
    })
  const strongPairs = analyzedKeyPairs.map((entry) => entry.mapping)
  if (strongPairs.length > 0) {
    return { pairs: strongPairs.slice(0, 2), unresolved: [], autoSelected: true }
  }

  const exactPairs = analyzeMatchHeaderPairs(sourceTable, targetTable, headerMappings)
    .filter((entry) => entry.score >= 90 && entry.sharedValueCount > 0)
    .sort((left, right) => {
      if (right.sharedValueCount !== left.sharedValueCount) {
        return right.sharedValueCount - left.sharedValueCount
      }
      return right.score - left.score
    })
    .map((entry) => entry.mapping)
  return { pairs: exactPairs.slice(0, 1), unresolved: [], autoSelected: true }
}

export function findBestTargetRowIndex(
  sourceRow: CellValue[],
  targetRows: CellValue[][],
  matchPairs: TableHeaderMapping[],
): {
  status: 'matched' | 'no_match' | 'ambiguous'
  rowIndex?: number
  score?: number
  candidateRowIndexes?: number[]
} {
  const matches: Array<{ rowIndex: number; score: number }> = []

  targetRows.forEach((targetRow, targetRowIndex) => {
    let totalScore = 0
    for (const pair of matchPairs) {
      const sourceValue = sourceRow[pair.sourceRelativeColumnIndex]
      const targetValue = targetRow[pair.targetRelativeColumnIndex]
      const score = getStrictValueMatchScore(targetValue, sourceValue)
      if (score === 0) {
        return
      }
      totalScore += score
    }

    matches.push({
      rowIndex: targetRowIndex,
      score: Math.round(totalScore / matchPairs.length),
    })
  })

  matches.sort((left, right) => right.score - left.score)
  const best = matches[0]
  if (!best) {
    return { status: 'no_match' }
  }

  const ambiguous = matches.filter((entry) => entry.score === best.score)
  if (ambiguous.length > 1) {
    return {
      status: 'ambiguous',
      score: best.score,
      candidateRowIndexes: ambiguous.map((entry) => entry.rowIndex),
    }
  }

  return {
    status: 'matched',
    rowIndex: best.rowIndex,
    score: best.score,
  }
}

export function buildTargetRowFromSource(
  sourceRow: CellValue[],
  sourceTable: TableDescriptor,
  targetTable: TableDescriptor,
  headerMappings: TableHeaderMapping[],
): { row: Array<CellValue | string>; derived: DerivedCellComputation[] } {
  const row = Array<CellValue | string>(targetTable.columnCount).fill(null)
  headerMappings.forEach((mapping) => {
    row[mapping.targetRelativeColumnIndex] = sourceRow[mapping.sourceRelativeColumnIndex]
  })

  const derived: DerivedCellComputation[] = []
  targetTable.columns.forEach((targetColumn) => {
    const derivedValue = deriveTargetColumnValue(sourceTable, sourceRow, targetColumn.header)
    if (!derivedValue) return

    row[targetColumn.relativeColumnIndex] = derivedValue.value
    derived.push({
      targetHeader: targetColumn.header,
      targetColumnLetter: targetColumn.columnLetter,
      value: derivedValue.value,
      reason: derivedValue.reason,
    })
  })

  return { row, derived }
}

function buildDistinctValueSet(table: TableDescriptor, relativeColumnIndex: number): Set<string> {
  const values = new Set<string>()

  table.dataRows.forEach((row) => {
    const normalized = normalizeText(row[relativeColumnIndex])
    if (normalized) {
      values.add(normalized)
    }
  })

  return values
}

export function analyzeMatchHeaderPairs(
  sourceTable: TableDescriptor,
  targetTable: TableDescriptor,
  pairs: TableHeaderMapping[],
): Array<MatchHeaderAnalysis & { mapping: TableHeaderMapping }> {
  return pairs.map((mapping) => {
    const sourceValues = buildDistinctValueSet(sourceTable, mapping.sourceRelativeColumnIndex)
    const targetValues = buildDistinctValueSet(targetTable, mapping.targetRelativeColumnIndex)
    const sharedValues = [...sourceValues].filter((value) => targetValues.has(value))

    return {
      mapping,
      sourceHeader: mapping.sourceHeader,
      targetHeader: mapping.targetHeader,
      score: mapping.score,
      sharedValueCount: sharedValues.length,
      sourceDistinctCount: sourceValues.size,
      targetDistinctCount: targetValues.size,
      sharedValueSamples: sharedValues.slice(0, 5),
    }
  })
}

function getBestSemanticColumn(
  table: TableDescriptor,
  queries: string[],
): TableColumnDescriptor | null {
  let bestColumn: TableColumnDescriptor | null = null
  let bestScore = 0

  table.columns.forEach((column) => {
    const score = queries.reduce((currentBest, query) => Math.max(currentBest, getHeaderMatchScore(column.header, query)), 0)
    if (score > bestScore) {
      bestScore = score
      bestColumn = column
    }
  })

  return bestScore >= 70 ? bestColumn : null
}

function deriveTargetColumnValue(
  sourceTable: TableDescriptor,
  sourceRow: CellValue[],
  targetHeader: string,
): { value: CellValue | string; reason: string } | null {
  const unitPriceColumn = getBestSemanticColumn(sourceTable, ['unit price', '单价', '综合单价', '总单价'])
  const quantityColumn = getBestSemanticColumn(sourceTable, ['quantity', 'qty', '订单数量', '数量', '总数量', '出货数量'])
  const totalColumn = getBestSemanticColumn(sourceTable, ['order total', 'total amount', 'amount', '总价', '总金额', '订单总额', '订单总价', '金额'])

  const targetLooksLikeUnitPrice = getHeaderMatchScore(targetHeader, 'unit price') >= 70
    || getHeaderMatchScore(targetHeader, '单价') >= 70
    || getHeaderMatchScore(targetHeader, '综合单价') >= 70
    || getHeaderMatchScore(targetHeader, '总单价') >= 70
  if (targetLooksLikeUnitPrice && quantityColumn && totalColumn) {
    const quantity = parseNumericValue(sourceRow[quantityColumn.relativeColumnIndex])
    const total = parseNumericValue(sourceRow[totalColumn.relativeColumnIndex])
    if (quantity && total !== null) {
      return {
        value: Number((total / quantity).toFixed(6)),
        reason: `Derived from ${totalColumn.header} / ${quantityColumn.header}`,
      }
    }
  }

  const targetLooksLikeAmount = getHeaderMatchScore(targetHeader, 'amount') >= 70
    || getHeaderMatchScore(targetHeader, 'total amount') >= 70
    || getHeaderMatchScore(targetHeader, '总价') >= 70
    || getHeaderMatchScore(targetHeader, '金额') >= 70
  if (targetLooksLikeAmount && unitPriceColumn && quantityColumn) {
    const unitPrice = parseNumericValue(sourceRow[unitPriceColumn.relativeColumnIndex])
    const quantity = parseNumericValue(sourceRow[quantityColumn.relativeColumnIndex])
    if (unitPrice !== null && quantity !== null) {
      return {
        value: Number((unitPrice * quantity).toFixed(6)),
        reason: `Derived from ${unitPriceColumn.header} * ${quantityColumn.header}`,
      }
    }
  }

  return null
}
