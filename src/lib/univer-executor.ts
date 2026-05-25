import {
  analyzeMatchHeaderPairs,
  buildHeaderMappings,
  buildTargetRowFromSource,
  findSheetByName,
  type DerivedCellComputation,
  findBestTargetRowIndex,
  isFormulaString,
  normalizeText,
  rangeToA1,
  readRecords,
  resolveSyncMatchHeaders,
  summarizeTable,
  tableColumnLetterToIndex,
} from './univer-schema'
import type { CellValue, TableDescriptor, WorkbookLike } from './univer-schema'
import { UniverSchemaCache } from './univer-cache'

export class UniverOperationExecutor {
  private readonly cache: UniverSchemaCache

  constructor(cache: UniverSchemaCache) {
    this.cache = cache
  }

  private valuesEquivalent(actual: unknown, expected: unknown): boolean {
    if (isFormulaString(expected)) {
      return String(actual ?? '') === String(expected)
    }

    if ((actual === null || actual === undefined || actual === '') && (expected === null || expected === undefined || expected === '')) {
      return true
    }

    return normalizeText(actual) === normalizeText(expected)
  }

  private readScalarVerification(
    workbook: WorkbookLike,
    sheetName: string,
    rowIndex: number,
    columnIndex: number,
    expectedValue: unknown,
  ) {
    const sheet = findSheetByName(workbook, sheetName)
    if (!sheet) {
      return {
        verified: false,
        error: `Sheet not found during verification: ${sheetName}`,
      }
    }

    const range = sheet.getRange(rowIndex, columnIndex)
    const actualValue = range.getValues()?.[0]?.[0] ?? null
    const actualFormula = range.getFormulas()?.[0]?.[0] ?? null
    const actualResolved = actualFormula || actualValue

    return {
      verified: this.valuesEquivalent(actualResolved, expectedValue),
      actual_value: actualValue,
      actual_formula: actualFormula,
      expected_value: expectedValue ?? null,
    }
  }

  private verifyRowValues(
    workbook: WorkbookLike,
    sheetName: string,
    startRow: number,
    startColumn: number,
    expectedRows: Array<Array<CellValue | string>>,
  ) {
    const sheet = findSheetByName(workbook, sheetName)
    if (!sheet) {
      return {
        verified: false,
        error: `Sheet not found during verification: ${sheetName}`,
        mismatches: [],
      }
    }

    const mismatches: Array<{
      row_number: number
      column_number: number
      expected_value: unknown
      actual_value: unknown
      actual_formula: string | null
    }> = []

    expectedRows.forEach((expectedRow, rowOffset) => {
      expectedRow.forEach((expectedValue, columnOffset) => {
        const range = sheet.getRange(startRow + rowOffset, startColumn + columnOffset)
        const actualValue = range.getValues()?.[0]?.[0] ?? null
        const actualFormula = range.getFormulas()?.[0]?.[0] ?? null
        const actualResolved = actualFormula || actualValue

        if (!this.valuesEquivalent(actualResolved, expectedValue)) {
          mismatches.push({
            row_number: startRow + rowOffset + 1,
            column_number: startColumn + columnOffset + 1,
            expected_value: expectedValue ?? null,
            actual_value: actualValue,
            actual_formula: actualFormula,
          })
        }
      })
    })

    return {
      verified: mismatches.length === 0,
      mismatch_count: mismatches.length,
      mismatches,
    }
  }

  private writeScalarToRange(
    range: {
      setFormula(value: string): void
      setValue(value: string | number | boolean): void
    },
    value: unknown,
  ): void {
    if (isFormulaString(value)) {
      range.setFormula(value)
      return
    }

    range.setValue(value as string | number | boolean)
  }

  private writeRowsToSheet(
    sheet: {
      getRange(row: number, column: number, numRows: number, numColumns: number): {
        setValues(values: (string | number | boolean)[][]): void
      }
      getRange(row: number, column: number): {
        setFormula(value: string): void
      }
    },
    startRow: number,
    startColumn: number,
    rows: Array<Array<CellValue | string>>,
  ): void {
    if (rows.length === 0) return

    const plainValues = rows.map((row) =>
      row.map((value) => (isFormulaString(value) ? null : (value as CellValue))),
    )
    sheet.getRange(startRow, startColumn, rows.length, rows[0].length).setValues(plainValues as (string | number | boolean)[][])

    rows.forEach((row, rowOffset) => {
      row.forEach((value, columnOffset) => {
        if (isFormulaString(value)) {
          sheet.getRange(startRow + rowOffset, startColumn + columnOffset).setFormula(value)
        }
      })
    })
  }

  setTableCellValue(
    workbook: WorkbookLike,
    location: {
      sheet_name: string
      target_cell: string
      current_value: CellValue
      current_formula: string | null
    },
    value: unknown,
  ) {
    const targetSheet = findSheetByName(workbook, location.sheet_name)
    if (!targetSheet) {
      return { error: `Sheet not found: ${location.sheet_name}` }
    }

    const columnLetter = location.target_cell.replace(/[0-9]/g, '')
    const rowNumber = Number(location.target_cell.replace(/[^0-9]/g, ''))
    const columnIndex = tableColumnLetterToIndex(columnLetter)
    const rowIndex = rowNumber - 1
    const targetRange = targetSheet.getRange(rowIndex, columnIndex)
    this.writeScalarToRange(targetRange, value)

    this.cache.invalidate(workbook)
    const verification = this.readScalarVerification(workbook, location.sheet_name, rowIndex, columnIndex, value)

    return {
      success: true,
      previous_value: location.current_value,
      previous_formula: location.current_formula,
      written_value: value ?? null,
      verification,
    }
  }

  appendTableRecords(
    workbook: WorkbookLike,
    targetTable: TableDescriptor,
    recordsInput: unknown,
  ) {
    const records = readRecords(recordsInput)
    if (records.length === 0) {
      return { error: 'records must contain at least one object row.' }
    }

    const targetSheet = findSheetByName(workbook, targetTable.sheetName)
    if (!targetSheet) {
      return { error: `Sheet not found: ${targetTable.sheetName}` }
    }

    const headerMappings = records.map((record, index) => {
      const resolved = Object.entries(record).map(([key, value]) => {
        const targetColumn = targetTable.columns.find((column) => column.normalizedHeader === key || column.header === key)
          ?? targetTable.columns.find((column) => column.normalizedHeader.includes(key) || key.includes(column.normalizedHeader))
        return targetColumn
          ? {
              source_key: key,
              target_header: targetColumn.header,
              target_column_letter: targetColumn.columnLetter,
              score: 100,
              value,
            }
          : {
              source_key: key,
              target_header: null,
              target_column_letter: null,
              score: 0,
              value,
            }
      })

      return { index, resolved }
    })

    const unresolvedKeys = headerMappings
      .flatMap((entry) => entry.resolved.filter((item) => !item.target_header).map((item) => item.source_key))

    const rowValues = headerMappings.map((entry) => {
      const row = Array<CellValue | string>(targetTable.columnCount).fill(null)
      entry.resolved.forEach((mapping) => {
        if (!mapping.target_header) return
        const targetColumn = targetTable.columns.find((column) => column.header === mapping.target_header)
        if (!targetColumn) return
        row[targetColumn.relativeColumnIndex] = isFormulaString(mapping.value)
          ? mapping.value
          : (mapping.value as CellValue)
      })
      return row
    })

    const appendStartRow = targetTable.headerRowIndex + 1 + targetTable.rowCount
    this.writeRowsToSheet(targetSheet, appendStartRow, targetTable.startColumn, rowValues)
    this.cache.invalidate(workbook)
    const verification = this.verifyRowValues(workbook, targetTable.sheetName, appendStartRow, targetTable.startColumn, rowValues)

    return {
      success: true,
      target_table: summarizeTable(targetTable),
      appended_rows: rowValues.length,
      write_range: rangeToA1(
        appendStartRow,
        targetTable.startColumn,
        appendStartRow + rowValues.length - 1,
        targetTable.startColumn + targetTable.columnCount - 1,
      ),
      unresolved_source_keys: [...new Set(unresolvedKeys)],
      header_mapping_preview: headerMappings,
      verification,
    }
  }

  syncTableToTable(
    workbook: WorkbookLike,
    sourceTable: TableDescriptor,
    targetTable: TableDescriptor,
    args: Record<string, unknown>,
  ) {
    const previewOnly = args.__preview_only === true
    const targetSheet = findSheetByName(workbook, targetTable.sheetName)
    if (!targetSheet) {
      return { error: `Sheet not found: ${targetTable.sheetName}` }
    }

    const headerMappings = buildHeaderMappings(sourceTable, targetTable)
    if (headerMappings.length === 0) {
      return { error: 'No semantic header mapping was found between the source and target tables.' }
    }

    const requestedMatchHeaders = Array.isArray(args.match_headers)
      ? args.match_headers
        .filter((value): value is string => typeof value === 'string' && value.trim().length > 0)
        .map((value) => value.trim())
      : []
    const matchResolution = resolveSyncMatchHeaders(sourceTable, targetTable, headerMappings, requestedMatchHeaders)
    if (requestedMatchHeaders.length > 0 && matchResolution.unresolved.length > 0) {
      return {
        error: `Some match headers could not be resolved: ${matchResolution.unresolved.join(', ')}`,
        header_mappings: headerMappings,
      }
    }

    const writeMode = args.write_mode === 'append' ? 'append' : 'upsert'
    const matchHeaderAnalysis = analyzeMatchHeaderPairs(sourceTable, targetTable, matchResolution.pairs)
    const validMatchPairs = matchHeaderAnalysis
      .filter((entry) => entry.sharedValueCount > 0)
      .map((entry) => entry.mapping)

    if (writeMode !== 'append' && validMatchPairs.length === 0) {
      return {
        error: requestedMatchHeaders.length > 0
          ? 'The selected match headers do not contain any shared values between the source and target tables. Please choose at least one true shared key column before syncing.'
          : 'No reliable shared key was found between the source and target tables. Please specify match_headers such as project, id, code, month, or another column that exists in both tables with shared values.',
        preview_only: previewOnly,
        source_table: summarizeTable(sourceTable),
        target_table: summarizeTable(targetTable),
        write_mode: args.write_mode === 'append' ? 'append' : 'upsert',
        auto_selected_match_headers: matchResolution.autoSelected,
        match_headers: matchResolution.pairs.map((pair) => ({
          source_header: pair.sourceHeader,
          target_header: pair.targetHeader,
          score: pair.score,
        })),
        match_header_analysis: matchHeaderAnalysis,
        header_mappings: headerMappings,
      }
    }

    const targetRows = targetTable.dataRows.map((row) => [...row])
    const appendedRows: Array<Array<CellValue | string>> = []
    const updatedRows: Array<{ row_number: number; match_score: number; matched_on: string[] }> = []
    const updatedRowPayloads: Array<{ rowIndex: number; values: Array<CellValue | string> }> = []
    const derived_cells: Array<{ source_row_number: number; target_row_number?: number; cells: DerivedCellComputation[] }> = []
    const skippedRows: Array<{ source_row_number: number; reason: string; record: Record<string, unknown> }> = []
    const unmatchedRows: Array<{ source_row_number: number; reason: string; record: Record<string, unknown> }> = []
    const ambiguousRows: Array<{ source_row_number: number; candidate_row_numbers: number[]; matched_on: string[]; record: Record<string, unknown> }> = []

    sourceTable.dataRows.forEach((sourceRow, sourceRowIndex) => {
      const sourceRecord = sourceTable.columns.reduce<Record<string, unknown>>((record, column) => {
        record[column.header] = sourceRow[column.relativeColumnIndex]
        return record
      }, {})
      const targetRowBuild = buildTargetRowFromSource(sourceRow, sourceTable, targetTable, headerMappings)
      const targetRowPayload = targetRowBuild.row

      if (writeMode === 'append') {
        appendedRows.push(targetRowPayload)
        targetRows.push(targetRowPayload.map((value) => isFormulaString(value) ? null : (value as CellValue)))
        if (targetRowBuild.derived.length > 0) {
          derived_cells.push({
            source_row_number: sourceTable.headerRowIndex + 2 + sourceRowIndex,
            cells: targetRowBuild.derived,
          })
        }
        return
      }

      const bestTargetRow = findBestTargetRowIndex(sourceRow, targetRows, validMatchPairs)
      if (bestTargetRow.status === 'no_match') {
        unmatchedRows.push({
          source_row_number: sourceTable.headerRowIndex + 2 + sourceRowIndex,
          reason: 'No target row shares the same key values.',
          record: sourceRecord,
        })
        return
      }

      if (bestTargetRow.status === 'ambiguous') {
        ambiguousRows.push({
          source_row_number: sourceTable.headerRowIndex + 2 + sourceRowIndex,
          candidate_row_numbers: (bestTargetRow.candidateRowIndexes ?? []).map((rowIndex) => targetTable.headerRowIndex + 2 + rowIndex),
          matched_on: validMatchPairs.map((pair) => `${pair.sourceHeader} -> ${pair.targetHeader}`),
          record: sourceRecord,
        })
        return
      }

      if ((bestTargetRow.score ?? 0) < 100 || bestTargetRow.rowIndex === undefined) {
        skippedRows.push({
          source_row_number: sourceTable.headerRowIndex + 2 + sourceRowIndex,
          reason: `Low comparison confidence (${bestTargetRow.score ?? 0}) for target row match.`,
          record: sourceRecord,
        })
        return
      }

      const targetDataRowIndex = bestTargetRow.rowIndex as number
      const targetSheetRowIndex = targetTable.headerRowIndex + 1 + targetDataRowIndex

      if (targetRowBuild.derived.length > 0) {
        derived_cells.push({
          source_row_number: sourceTable.headerRowIndex + 2 + sourceRowIndex,
          target_row_number: targetSheetRowIndex + 1,
          cells: targetRowBuild.derived,
        })
      }

      if (!previewOnly) {
        headerMappings.forEach((mapping) => {
          const value = targetRowPayload[mapping.targetRelativeColumnIndex]
          const range = targetSheet.getRange(targetSheetRowIndex, targetTable.startColumn + mapping.targetRelativeColumnIndex)
          this.writeScalarToRange(range, value)
          targetRows[targetDataRowIndex][mapping.targetRelativeColumnIndex] = isFormulaString(value) ? null : (value as CellValue)
        })
        updatedRowPayloads.push({
          rowIndex: targetSheetRowIndex,
          values: targetRowPayload,
        })
      }

      updatedRows.push({
        row_number: targetSheetRowIndex + 1,
        match_score: bestTargetRow.score ?? 100,
        matched_on: validMatchPairs.map((pair) => `${pair.sourceHeader} -> ${pair.targetHeader}`),
      })
    })

    const appendRowIndex = targetTable.headerRowIndex + 1 + targetTable.rowCount
    if (!previewOnly) {
      this.writeRowsToSheet(targetSheet, appendRowIndex, targetTable.startColumn, appendedRows)
      this.cache.invalidate(workbook)
    }

    const updateVerifications = !previewOnly
      ? updatedRowPayloads.map((entry) => this.verifyRowValues(
        workbook,
        targetTable.sheetName,
        entry.rowIndex,
        targetTable.startColumn,
        [entry.values],
      ))
      : []
    const appendVerification = !previewOnly && appendedRows.length > 0
      ? this.verifyRowValues(workbook, targetTable.sheetName, appendRowIndex, targetTable.startColumn, appendedRows)
      : null
    const verification = !previewOnly
      ? {
        verified: updateVerifications.every((entry) => entry.verified) && (appendVerification?.verified ?? true),
        updated_row_checks: updateVerifications,
        appended_row_check: appendVerification,
      }
      : null

    return {
      success: true,
      preview_only: previewOnly,
      source_table: summarizeTable(sourceTable),
      target_table: summarizeTable(targetTable),
      write_mode: writeMode,
      auto_selected_match_headers: matchResolution.autoSelected,
      match_headers: validMatchPairs.map((pair) => ({
        source_header: pair.sourceHeader,
        target_header: pair.targetHeader,
        score: pair.score,
      })),
      match_header_analysis: matchHeaderAnalysis,
      header_mappings: headerMappings,
      appended_rows: appendedRows.length,
      appended_write_range: appendedRows.length > 0
        ? rangeToA1(
            appendRowIndex,
            targetTable.startColumn,
            appendRowIndex + appendedRows.length - 1,
            targetTable.startColumn + targetTable.columnCount - 1,
          )
        : null,
      updated_rows: updatedRows,
      skipped_rows: skippedRows,
      unmatched_rows: unmatchedRows,
      ambiguous_rows: ambiguousRows,
      derived_cells,
      verification,
    }
  }
}
