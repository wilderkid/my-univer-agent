import { getUniverAPI } from './univer-ref'
import { UniverSchemaCache } from './univer-cache'
import { UniverOperationExecutor } from './univer-executor'
import { createPlanId, UniverPlanStore } from './univer-plan-store'
import {
  colIndexToLetter,
  CONTEXT_PREVIEW_MAX_COLUMNS,
  CONTEXT_PREVIEW_MAX_ROWS,
  findSheetByName,
  findRowsByConditions,
  isFormulaString,
  locateTargetCell,
  rangeToA1,
  resolveSnapshotSheetName,
  resolveTable,
  resolveTableByPrefix,
  summarizeTable,
  tableColumnLetterToIndex,
} from './univer-schema'
import type { CellValue, SheetLike, TableDescriptor, WorkbookLike } from './univer-schema'

const API_BASE = import.meta.env.VITE_APP_API_BASE ?? ''

export interface UniverAgentConfig {
  apiKey: string
  model: string
  baseURL: string
}

export interface AgentToolExecution {
  toolName: string
  args: Record<string, unknown>
  result: unknown
}

export interface AgentPlanPreview {
  planId: string
  type: string
  summary: string
  result: Record<string, unknown>
}

export interface AgentExecutionResult {
  message: string
  finalMessage: string | null
  tools: AgentToolExecution[]
  plans: AgentPlanPreview[]
}

export interface AgentApplyPlanResult {
  message: string
  result: unknown
}

export interface AgentSnapshotSummary {
  snapshotId: string
  createdAt: number
  sheetCount: number
  cellCount: number
}

export interface AgentRestoreSnapshotResult {
  success: boolean
  snapshotId: string
  restoredSheets?: number
  restoredCells?: number
  message: string
  error?: string
}

export interface AgentCellDiff {
  sheetName: string
  cell: string
  beforeValue: unknown
  afterValue: unknown
  beforeFormula: string | null
  afterFormula: string | null
  kind: 'value' | 'formula' | 'value_formula' | 'format'
}

export interface AgentSnapshotDiff {
  success: boolean
  beforeSnapshotId: string
  afterSnapshotId: string
  changedCellCount: number
  valueChangeCount: number
  formulaChangeCount: number
  formatChangeCount: number
  addedSheetNames: string[]
  removedSheetNames: string[]
  changes: AgentCellDiff[]
  truncated: boolean
  message: string
  error?: string
}

interface ToolCall {
  id: string
  type: 'function'
  function: { name: string; arguments: string }
}

type Message =
  | { role: 'system'; content: string }
  | { role: 'user'; content: string }
  | { role: 'assistant'; content: string | null; tool_calls?: ToolCall[] }
  | { role: 'tool'; tool_call_id: string; content: string }

type LLMRequestMessage =
  | { role: 'system'; content: string }
  | { role: 'user'; content: string }
  | { role: 'assistant'; content: string; tool_calls?: ToolCall[] }
  | { role: 'tool'; tool_call_id: string; content: string }

interface LLMToolDefinition {
  type: 'function'
  function: {
    name: string
    description: string
    parameters: Record<string, unknown>
  }
}

type LLMPayloadMode = 'standard' | 'compact' | 'minimal'

interface ExecutionPolicy {
  allowDataTools: boolean
  allowFormatTools: boolean
  allowStructureTools: boolean
  allowClearTools: boolean
  intentSummary: string
}

interface SheetSnapshot {
  sheetId: string
  sheetName: string
  sheetIndex: number
  usedRange: string | null
  rowCount: number
  columnCount: number
  values: unknown[][]
  formulas: string[][]
  cellData: Array<Array<Record<string, unknown> | null>>
  columnWidths: number[]
  rowHeights: number[]
  mergedRanges: string[]
}

interface WorkbookSnapshot {
  id: string
  createdAt: number
  label: string
  sheets: SheetSnapshot[]
}

const TOOLS = [
  {
    type: 'function',
    function: {
      name: 'get_spreadsheet_info',
      description: 'Get workbook structure, active sheet, active cell, selected ranges, semantic table summaries, used range, and previews of current data before taking action.',
      parameters: { type: 'object', properties: {}, required: [] },
    },
  },
  {
    type: 'function',
    function: {
      name: 'list_tables',
      description: 'List detected semantic tables in the workbook or a specific sheet, including headers, preview rows, and likely key columns.',
      parameters: {
        type: 'object',
        properties: {
          sheet_name: { type: 'string', description: 'Optional sheet name. Defaults to the active sheet.' },
        },
        required: [],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'describe_table',
      description: 'Inspect one semantic table in detail. Use this before cross-sheet mapping or bulk writes.',
      parameters: {
        type: 'object',
        properties: {
          sheet_name: { type: 'string', description: 'Optional sheet name. Defaults to the active sheet.' },
          table_range: { type: 'string', description: 'Optional detected table range such as A1:F20.' },
        },
        required: [],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'find_rows_by_conditions',
      description: 'Find rows inside a semantic table by matching row conditions against table headers. Example row_conditions: {"Name":"Alice","Month":"Jan"}.',
      parameters: {
        type: 'object',
        properties: {
          sheet_name: { type: 'string', description: 'Optional sheet name. Defaults to the active sheet.' },
          table_range: { type: 'string', description: 'Optional detected table range such as A1:F20.' },
          row_conditions: {
            type: 'object',
            description: 'Map of header names to expected values.',
            additionalProperties: true,
          },
          limit: { type: 'number', description: 'Maximum number of matches to return. Defaults to 5.' },
        },
        required: ['row_conditions'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'locate_target_cell',
      description: 'Locate one exact target cell inside a semantic table by combining a target column header with row conditions. Use this before editing existing structured data.',
      parameters: {
        type: 'object',
        properties: {
          sheet_name: { type: 'string', description: 'Optional sheet name. Defaults to the active sheet.' },
          table_range: { type: 'string', description: 'Optional detected table range such as A1:F20.' },
          column_header: { type: 'string', description: 'The target column header, such as Score or Amount.' },
          row_conditions: {
            type: 'object',
            description: 'Map of header names to expected values used to identify the row.',
            additionalProperties: true,
          },
        },
        required: ['column_header', 'row_conditions'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'preview_set_table_cell_value',
      description: 'Create a preview plan for a semantic table cell update without writing anything. Use this before applying important edits.',
      parameters: {
        type: 'object',
        properties: {
          sheet_name: { type: 'string' },
          table_range: { type: 'string' },
          column_header: { type: 'string' },
          row_conditions: {
            type: 'object',
            additionalProperties: true,
          },
          value: { description: 'The value that will be written when the plan is applied.' },
        },
        required: ['column_header', 'row_conditions', 'value'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_table_cell_value',
      description: 'Safely write one value into an existing semantic table by target column header and row conditions. This tool locates and validates the target before writing.',
      parameters: {
        type: 'object',
        properties: {
          sheet_name: { type: 'string', description: 'Optional sheet name. Defaults to the active sheet.' },
          table_range: { type: 'string', description: 'Optional detected table range such as A1:F20.' },
          column_header: { type: 'string', description: 'The target column header, such as Score or Amount.' },
          row_conditions: {
            type: 'object',
            description: 'Map of header names to expected values used to identify the row.',
            additionalProperties: true,
          },
          value: { description: 'Value to write. Strings starting with = are formulas.' },
        },
        required: ['column_header', 'row_conditions', 'value'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'preview_sync_table_to_table',
      description: 'Create a preview plan for syncing a source table into a target table without writing anything. Use this before multi-table reorganization.',
      parameters: {
        type: 'object',
        properties: {
          source_sheet_name: { type: 'string' },
          source_table_range: { type: 'string' },
          target_sheet_name: { type: 'string' },
          target_table_range: { type: 'string' },
          match_headers: {
            type: 'array',
            items: { type: 'string' },
          },
          write_mode: {
            type: 'string',
            enum: ['append', 'upsert'],
          },
        },
        required: [],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'sync_table_to_table',
      description: 'Safely copy rows from a source semantic table into a target semantic table by matching headers and comparing key fields. This preserves row-level relationships such as project and amount.',
      parameters: {
        type: 'object',
        properties: {
          source_sheet_name: { type: 'string', description: 'Optional source sheet name. Defaults to the active sheet if the source table is omitted.' },
          source_table_range: { type: 'string', description: 'Optional source table range such as A1:F20.' },
          target_sheet_name: { type: 'string', description: 'Optional target sheet name. Defaults to the active sheet if the target table is omitted.' },
          target_table_range: { type: 'string', description: 'Optional target table range such as A1:F20.' },
          match_headers: {
            type: 'array',
            description: 'Optional semantic headers used to compare source rows against target rows, such as ["Project"] or ["Project", "Month"].',
            items: { type: 'string' },
          },
          write_mode: {
            type: 'string',
            enum: ['append', 'upsert'],
            description: 'append always adds new rows. upsert compares key fields and updates matched rows or appends new ones. Defaults to upsert.',
          },
        },
        required: [],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'apply_plan',
      description: 'Apply a previously previewed operation plan by plan_id. Use this only after inspecting the preview.',
      parameters: {
        type: 'object',
        properties: {
          plan_id: { type: 'string', description: 'The plan id returned by a preview tool.' },
        },
        required: ['plan_id'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'append_table_records',
      description: 'Append already-verified structured records into an existing semantic table by matching record keys to target headers. Prefer sync_table_to_table when moving rows directly from one table to another.',
      parameters: {
        type: 'object',
        properties: {
          sheet_name: { type: 'string', description: 'Optional target sheet name. Defaults to the active sheet.' },
          table_range: { type: 'string', description: 'Optional detected target table range such as A1:F20.' },
          records: {
            type: 'array',
            description: 'Array of objects. Each object is a row keyed by semantic header name.',
            items: {
              type: 'object',
              additionalProperties: true,
            },
          },
        },
        required: ['records'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_cell_value',
      description: 'Write one scalar value or one formula to an explicit A1 range. Use this only for explicit coordinates or building brand-new layouts, not for semantic table edits.',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string', description: 'A1 notation such as A1 or B2:D5' },
          value: { description: 'Value to write. May be string, number, boolean, or formula text.' },
          sheet_name: { type: 'string', description: 'Optional sheet name. Defaults to the active sheet.' },
        },
        required: ['range', 'value'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_range_values',
      description: 'Write a 2D array into an explicit target range. Use this only when the target A1 range is already known exactly. Formula cells must start with =.',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string', description: 'A1 notation such as A1:D4' },
          values: {
            type: 'array',
            description: 'A 2D array matching the target range size.',
            items: { type: 'array', items: {} },
          },
          sheet_name: { type: 'string', description: 'Optional sheet name. Defaults to the active sheet.' },
        },
        required: ['range', 'values'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'get_range_values',
      description: 'Read values and formulas from an exact range. Use this before modifying existing data if anything is ambiguous.',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string', description: 'A1 notation' },
          sheet_name: { type: 'string', description: 'Optional sheet name. Defaults to the active sheet.' },
        },
        required: ['range'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'format_range',
      description: 'Apply non-destructive formatting such as bold, italic, font size, colors, underline, alignment, wrap strategy, and borders. This never changes cell values.',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string' },
          bold: { type: 'boolean' },
          italic: { type: 'boolean' },
          font_family: { type: 'string', description: 'Font family name, such as Arial, Calibri, Microsoft YaHei, or SimSun.' },
          font_size: { type: 'number' },
          font_color: { type: 'string', description: 'CSS color such as #ff0000 or red' },
          background: { type: 'string', description: 'CSS background color' },
          underline: { type: 'boolean' },
          horizontal_alignment: { type: 'string', enum: ['left', 'center', 'right'] },
          vertical_alignment: { type: 'string', enum: ['top', 'middle', 'bottom'] },
          text_rotation: { type: 'number', description: 'Text rotation angle in degrees. Typical values are 0, 45, 90, or -45.' },
          wrap_strategy: { type: 'string', enum: ['overflow', 'wrap', 'clip'], description: 'Text display strategy. Use wrap for long text when appropriate.' },
          number_format: { type: 'string', description: 'Spreadsheet number format pattern, such as #,##0.000 or $#,##0.000. Use this for display precision without changing values.' },
          border_type: { type: 'string', enum: ['top', 'bottom', 'left', 'right', 'all', 'outside', 'inside', 'horizontal', 'vertical'] },
          border_style: { type: 'string', enum: ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'none'] },
          border_color: { type: 'string', description: 'CSS border color. Defaults to #000000.' },
          sheet_name: { type: 'string' },
        },
        required: ['range'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_freeze',
      description: 'Freeze or unfreeze rows and columns on a worksheet. Use 0 for frozen_rows or frozen_columns to unfreeze that direction.',
      parameters: {
        type: 'object',
        properties: {
          frozen_rows: { type: 'number', description: 'Number of top rows to freeze. Defaults to current/0 if omitted.' },
          frozen_columns: { type: 'number', description: 'Number of left columns to freeze. Defaults to current/0 if omitted.' },
          sheet_name: { type: 'string' },
        },
        required: [],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'hide_rows',
      description: 'Hide one or more consecutive rows without deleting data.',
      parameters: {
        type: 'object',
        properties: {
          row_index: { type: 'number', description: 'Zero-based start row index, e.g. 0 for row 1.' },
          count: { type: 'number', description: 'Number of rows to hide. Defaults to 1.' },
          sheet_name: { type: 'string' },
        },
        required: ['row_index'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'show_rows',
      description: 'Show one or more previously hidden consecutive rows.',
      parameters: {
        type: 'object',
        properties: {
          row_index: { type: 'number', description: 'Zero-based start row index, e.g. 0 for row 1.' },
          count: { type: 'number', description: 'Number of rows to show. Defaults to 1.' },
          sheet_name: { type: 'string' },
        },
        required: ['row_index'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'hide_columns',
      description: 'Hide one or more consecutive columns without deleting data.',
      parameters: {
        type: 'object',
        properties: {
          column_index: { type: 'number', description: 'Zero-based start column index, e.g. 0 for A.' },
          count: { type: 'number', description: 'Number of columns to hide. Defaults to 1.' },
          sheet_name: { type: 'string' },
        },
        required: ['column_index'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'show_columns',
      description: 'Show one or more previously hidden consecutive columns.',
      parameters: {
        type: 'object',
        properties: {
          column_index: { type: 'number', description: 'Zero-based start column index, e.g. 0 for A.' },
          count: { type: 'number', description: 'Number of columns to show. Defaults to 1.' },
          sheet_name: { type: 'string' },
        },
        required: ['column_index'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_gridlines',
      description: 'Show or hide worksheet gridlines. This is a display-only operation and does not change cell data.',
      parameters: {
        type: 'object',
        properties: {
          hidden: { type: 'boolean', description: 'true hides gridlines; false shows gridlines.' },
          sheet_name: { type: 'string' },
        },
        required: ['hidden'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_sheet_tab_color',
      description: 'Set the worksheet tab color when supported by the current Univer build.',
      parameters: {
        type: 'object',
        properties: {
          color: { type: 'string', description: 'CSS color such as #22c55e or red.' },
          sheet_name: { type: 'string' },
        },
        required: ['color'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_column_width',
      description: 'Set a spreadsheet column width in pixels. Use this for display/layout changes instead of merging cells.',
      parameters: {
        type: 'object',
        properties: {
          column_index: { type: 'number', description: 'Zero-based column index, e.g. 0 for A and 1 for B.' },
          width: { type: 'number', description: 'Column width in pixels. Typical values are 80-360.' },
          sheet_name: { type: 'string' },
        },
        required: ['column_index', 'width'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_row_height',
      description: 'Set one or more spreadsheet row heights in pixels. Use this when the user asks to unify or adjust row height. If row_index is omitted, the active range or used range rows are used.',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string', description: 'Optional A1 range whose rows should be changed. If omitted, the active range or used range is used.' },
          row_index: { type: 'number', description: 'Zero-based start row index, e.g. 0 for row 1.' },
          count: { type: 'number', description: 'Number of consecutive rows to set. Defaults to 1.' },
          height: { type: 'number', description: 'Row height in pixels. Typical values are 22-80.' },
          forced: { type: 'boolean', description: 'When true, force the exact height even if content is taller. Defaults to true.' },
          sheet_name: { type: 'string' },
        },
        required: ['height'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'auto_resize_columns',
      description: 'Resize columns in a range based on visible text length. This is safe for long text and does not merge cells or change values.',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string', description: 'A1 range whose columns should be auto-sized. If omitted, the active range is used.' },
          min_width: { type: 'number', description: 'Minimum width in pixels. Defaults to 80.' },
          max_width: { type: 'number', description: 'Maximum width in pixels. Defaults to 420.' },
          padding: { type: 'number', description: 'Extra width in pixels. Defaults to 28.' },
          sheet_name: { type: 'string' },
        },
        required: [],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'auto_resize_rows',
      description: 'Resize rows in a range based on text length and wrapping. This is safe for long text and does not merge cells or change values.',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string', description: 'A1 range whose rows should be auto-sized. If omitted, the active range is used.' },
          min_height: { type: 'number', description: 'Minimum height in pixels. Defaults to 22.' },
          max_height: { type: 'number', description: 'Maximum height in pixels. Defaults to 180.' },
          base_line_height: { type: 'number', description: 'Estimated line height in pixels. Defaults to 20.' },
          sheet_name: { type: 'string' },
        },
        required: [],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'clear_range',
      description: 'Clear contents and/or formatting in a range.',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string' },
          contents_only: { type: 'boolean', description: 'Clear only contents and keep formatting.' },
          format_only: { type: 'boolean', description: 'Clear only formatting and keep contents.' },
          sheet_name: { type: 'string' },
        },
        required: ['range'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'insert_sheet',
      description: 'Create a new worksheet.',
      parameters: {
        type: 'object',
        properties: {
          name: { type: 'string', description: 'New sheet name' },
        },
        required: ['name'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'rename_sheet',
      description: 'Rename a worksheet. If old_name is omitted, rename the active sheet.',
      parameters: {
        type: 'object',
        properties: {
          new_name: { type: 'string' },
          old_name: { type: 'string', description: 'Optional existing sheet name' },
        },
        required: ['new_name'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'insert_rows',
      description: 'Insert rows before the given zero-based row index.',
      parameters: {
        type: 'object',
        properties: {
          row_index: { type: 'number', description: 'Zero-based row index' },
          count: { type: 'number', description: 'Row count to insert. Defaults to 1.' },
          sheet_name: { type: 'string' },
        },
        required: ['row_index'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'delete_rows',
      description: 'Delete rows starting from the given zero-based row index.',
      parameters: {
        type: 'object',
        properties: {
          row_index: { type: 'number' },
          count: { type: 'number', description: 'Row count to delete. Defaults to 1.' },
          sheet_name: { type: 'string' },
        },
        required: ['row_index'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'insert_columns',
      description: 'Insert columns before the given zero-based column index.',
      parameters: {
        type: 'object',
        properties: {
          column_index: { type: 'number' },
          count: { type: 'number', description: 'Column count to insert. Defaults to 1.' },
          sheet_name: { type: 'string' },
        },
        required: ['column_index'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'delete_columns',
      description: 'Delete columns starting from the given zero-based column index.',
      parameters: {
        type: 'object',
        properties: {
          column_index: { type: 'number' },
          count: { type: 'number', description: 'Column count to delete. Defaults to 1.' },
          sheet_name: { type: 'string' },
        },
        required: ['column_index'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'merge_cells',
      description: 'Merge or unmerge a range when the user explicitly asks to merge or unmerge cells. Merging may hide non-top-left cell contents, so use it only for explicit structure/layout requests.',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string' },
          unmerge: { type: 'boolean', description: 'When true, unmerge instead of merge.' },
          sheet_name: { type: 'string' },
        },
        required: ['range'],
      },
    },
  },
] as const

type ToolDefinition = (typeof TOOLS)[number]
type ToolName = ToolDefinition['function']['name']

const READ_TOOL_NAMES = new Set<ToolName>([
  'get_spreadsheet_info',
  'list_tables',
  'describe_table',
  'find_rows_by_conditions',
  'locate_target_cell',
  'get_range_values',
])

const DATA_TOOL_NAMES = new Set<ToolName>([
  'preview_set_table_cell_value',
  'set_table_cell_value',
  'preview_sync_table_to_table',
  'sync_table_to_table',
  'apply_plan',
  'append_table_records',
  'set_cell_value',
  'set_range_values',
])

const FORMAT_TOOL_NAMES = new Set<ToolName>([
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
])

const STRUCTURE_TOOL_NAMES = new Set<ToolName>([
  'insert_sheet',
  'rename_sheet',
  'insert_rows',
  'delete_rows',
  'insert_columns',
  'delete_columns',
  'merge_cells',
])

const MINIMAL_COMPATIBLE_TOOL_NAMES = new Set<ToolName>([
  'get_spreadsheet_info',
  'list_tables',
  'describe_table',
  'find_rows_by_conditions',
  'locate_target_cell',
  'preview_set_table_cell_value',
  'set_table_cell_value',
  'preview_sync_table_to_table',
  'sync_table_to_table',
  'apply_plan',
  'append_table_records',
  'set_cell_value',
  'set_range_values',
  'get_range_values',
])

const SYSTEM_PROMPT = [
  'You are an AI assistant that controls a Univer spreadsheet.',
  'You must use tools to perform actions, not just describe them.',
  'Row and column indexes used by row_index and column_index are zero-based.',
  'A1 notation is one-based, such as A1, B3, and C1:D5.',
  'Before changing existing data, first inspect the workbook state with get_spreadsheet_info or list_tables.',
  'For structured table edits, prefer describe_table, find_rows_by_conditions, locate_target_cell, preview_set_table_cell_value, preview_sync_table_to_table, apply_plan, and append_table_records.',
  'Use set_cell_value and set_range_values only when the target A1 coordinates are explicit or when building a new empty layout.',
  'When the user refers to current cells, selected cells, this area, this table, or nearby data, pay close attention to the active sheet and selected ranges.',
  'Do not guess table shapes or target positions. If a semantic location is ambiguous, inspect more and avoid writing.',
  'Preview first, then apply. Prefer preview_set_table_cell_value and preview_sync_table_to_table before any important write.',
  'For bulk reorganization across sheets, prefer preview_sync_table_to_table so the system preserves each source row while matching headers and key fields.',
  'If the target needs one derived business value such as unit price while the source contains total amount and quantity, prefer deriving the value from total amount divided by quantity instead of guessing among multiple candidate price columns.',
  'Prefer precise minimal writes. Do not overwrite nearby cells unless the user clearly asked for that.',
  'For formatting-only requests, use format_range, set_column_width, set_row_height, auto_resize_columns, auto_resize_rows, set_freeze, hide_rows, show_rows, hide_columns, show_columns, set_gridlines, and set_sheet_tab_color. Never change cell values, clear contents, insert/delete rows or columns, or merge cells just to improve appearance.',
  'For row-height requests such as "统一行高", "调整行高", "设置行高", or "row height", call set_row_height. If the user does not specify rows, omit row_index and use the selected range or used range. Use auto_resize_rows only when the user asks for automatic/adaptive height based on content.',
  'When the user asks for font family, vertical alignment, text rotation, gridlines, frozen panes, hidden rows/columns, or sheet tab color, use the dedicated formatting/display tools instead of changing data.',
  'Do not use merged cells as an automatic workaround for long text. Use column width and wrap strategy unless the user explicitly asks to merge cells.',
  'Never ask the user for confirmation inside the final answer. If a request is allowed by tools and policy, execute it. If it is unsafe or ambiguous, stop and state the concrete reason.',
  'If the user asks to change displayed decimal places, use number_format in format_range. Do not convert numeric values to text and do not clear existing values.',
  'The user may speak Chinese. Reply with a short Chinese result summary after the actions finish.',
].join('\n')

export class UniverAgent {
  private static readonly REQUEST_TIMEOUT_MS = 60_000
  private static readonly MAX_TOOL_RESULT_ARRAY_ITEMS = 12
  private static readonly MAX_TOOL_RESULT_STRING_LENGTH = 600
  private static readonly MAX_TOOL_RESULT_JSON_LENGTH = 16_000

  private readonly config: UniverAgentConfig
  private readonly schemaCache = new UniverSchemaCache()
  private readonly executor = new UniverOperationExecutor(this.schemaCache)
  private readonly planStore = new UniverPlanStore()
  private readonly snapshots = new Map<string, WorkbookSnapshot>()
  private activeController: AbortController | null = null
  private activeAbortReason: string | null = null
  private activePolicy: ExecutionPolicy | null = null

  constructor(config: UniverAgentConfig) {
    this.config = config
  }

  cancel(reason = 'Execution canceled'): void {
    this.activeAbortReason = reason
    this.activeController?.abort()
  }

  createWorkbookSnapshot(label = 'Before AI operation'): AgentSnapshotSummary {
    const api = getUniverAPI()
    const workbook = api?.getActiveWorkbook()
    if (!workbook) throw new Error('No active workbook')

    const typedWorkbook = workbook as unknown as WorkbookLike
    const id = `snapshot-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`
    const sheets = typedWorkbook.getSheets().map((sheet, sheetIndex) => this.captureSheetSnapshot(sheet, sheetIndex))
    const snapshot: WorkbookSnapshot = {
      id,
      createdAt: Date.now(),
      label,
      sheets,
    }

    this.snapshots.set(id, snapshot)

    // Keep recent history bounded; UI only displays the newest execution records.
    if (this.snapshots.size > 30) {
      const oldestId = [...this.snapshots.keys()][0]
      if (oldestId) this.snapshots.delete(oldestId)
    }

    return {
      snapshotId: snapshot.id,
      createdAt: snapshot.createdAt,
      sheetCount: snapshot.sheets.length,
      cellCount: snapshot.sheets.reduce((sum, sheet) => sum + sheet.rowCount * sheet.columnCount, 0),
    }
  }

  restoreWorkbookSnapshot(snapshotId: string): AgentRestoreSnapshotResult {
    const snapshot = this.snapshots.get(snapshotId)
    if (!snapshot) {
      return {
        success: false,
        snapshotId,
        message: '未找到执行前快照，无法恢复。',
        error: 'Snapshot not found.',
      }
    }

    const api = getUniverAPI()
    const workbook = api?.getActiveWorkbook()
    if (!workbook) {
      return {
        success: false,
        snapshotId,
        message: 'Univer API 未初始化，无法恢复快照。',
        error: 'No active workbook.',
      }
    }

    const typedWorkbook = workbook as unknown as WorkbookLike

    try {
      this.syncWorkbookSheetsToSnapshot(typedWorkbook, snapshot)

      let restoredCells = 0
      snapshot.sheets
        .slice()
        .sort((left, right) => left.sheetIndex - right.sheetIndex)
        .forEach((sheetSnapshot) => {
        const sheet = typedWorkbook.getSheetByName(sheetSnapshot.sheetName) ?? typedWorkbook.insertSheet(sheetSnapshot.sheetName)
        this.restoreSheetSnapshot(sheet, sheetSnapshot)
        restoredCells += sheetSnapshot.rowCount * sheetSnapshot.columnCount
      })

      this.schemaCache.invalidate(typedWorkbook)

      return {
        success: true,
        snapshotId,
        restoredSheets: snapshot.sheets.length,
        restoredCells,
        message: `已恢复到本次 AI 执行前：${snapshot.sheets.length} 个工作表，${restoredCells} 个单元格。`,
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      return {
        success: false,
        snapshotId,
        message: `恢复快照失败：${message}`,
        error: message,
      }
    }
  }

  diffWorkbookSnapshots(beforeSnapshotId: string, afterSnapshotId: string, limit = 80): AgentSnapshotDiff {
    const beforeSnapshot = this.snapshots.get(beforeSnapshotId)
    const afterSnapshot = this.snapshots.get(afterSnapshotId)

    if (!beforeSnapshot || !afterSnapshot) {
      return {
        success: false,
        beforeSnapshotId,
        afterSnapshotId,
        changedCellCount: 0,
        valueChangeCount: 0,
        formulaChangeCount: 0,
        formatChangeCount: 0,
        addedSheetNames: [],
        removedSheetNames: [],
        changes: [],
        truncated: false,
        message: '未找到快照，无法生成差异。',
        error: 'Snapshot not found.',
      }
    }

    const beforeSheets = new Map(beforeSnapshot.sheets.map((sheet) => [sheet.sheetName, sheet]))
    const afterSheets = new Map(afterSnapshot.sheets.map((sheet) => [sheet.sheetName, sheet]))
    const sheetNames = [...new Set([...beforeSheets.keys(), ...afterSheets.keys()])]
    const addedSheetNames = sheetNames.filter((sheetName) => !beforeSheets.has(sheetName))
    const removedSheetNames = sheetNames.filter((sheetName) => !afterSheets.has(sheetName))
    const changes: AgentCellDiff[] = []
    let changedCellCount = 0
    let valueChangeCount = 0
    let formulaChangeCount = 0
    let formatChangeCount = 0

    sheetNames.forEach((sheetName) => {
      const beforeSheet = beforeSheets.get(sheetName)
      const afterSheet = afterSheets.get(sheetName)
      const maxRows = Math.max(beforeSheet?.rowCount ?? 0, afterSheet?.rowCount ?? 0)
      const maxColumns = Math.max(beforeSheet?.columnCount ?? 0, afterSheet?.columnCount ?? 0)

      for (let rowIndex = 0; rowIndex < maxRows; rowIndex += 1) {
        for (let columnIndex = 0; columnIndex < maxColumns; columnIndex += 1) {
          const beforeCell = this.readSnapshotCell(beforeSheet, rowIndex, columnIndex)
          const afterCell = this.readSnapshotCell(afterSheet, rowIndex, columnIndex)
          const valueChanged = !this.snapshotValuesEqual(beforeCell.value, afterCell.value)
          const formulaChanged = (beforeCell.formula ?? '') !== (afterCell.formula ?? '')
          const formatChanged = !this.snapshotValuesEqual(beforeCell.style, afterCell.style)

          if (!valueChanged && !formulaChanged && !formatChanged) continue

          changedCellCount += 1
          if (valueChanged) valueChangeCount += 1
          if (formulaChanged) formulaChangeCount += 1
          if (formatChanged) formatChangeCount += 1

          if (changes.length < limit) {
            changes.push({
              sheetName,
              cell: `${colIndexToLetter(columnIndex)}${rowIndex + 1}`,
              beforeValue: beforeCell.value ?? null,
              afterValue: afterCell.value ?? null,
              beforeFormula: beforeCell.formula || null,
              afterFormula: afterCell.formula || null,
              kind: formulaChanged && valueChanged
                ? 'value_formula'
                : formulaChanged
                  ? 'formula'
                  : valueChanged
                    ? 'value'
                    : 'format',
            })
          }
        }
      }
    })

    return {
      success: true,
      beforeSnapshotId,
      afterSnapshotId,
      changedCellCount,
      valueChangeCount,
      formulaChangeCount,
      formatChangeCount,
      addedSheetNames,
      removedSheetNames,
      changes,
      truncated: changes.length < changedCellCount,
      message: changedCellCount > 0
        ? `检测到 ${changedCellCount} 项变化。`
        : '未检测到单元格数据、公式或格式变化。',
    }
  }

  async execute(instruction: string): Promise<string> {
    const result = await this.executeDetailed(instruction)
    return result.message
  }

  async executeDetailed(instruction: string): Promise<AgentExecutionResult> {
    const api = getUniverAPI()
    if (!api) throw new Error('Univer API is not initialized')

    this.cancel('Superseded by a new execution')
    const controller = new AbortController()
    this.activeController = controller
    this.activeAbortReason = null

    const context = this.getContext()
    const policy = this.createExecutionPolicy(instruction)
    this.activePolicy = policy
    const messages: Message[] = [
      { role: 'system', content: `${SYSTEM_PROMPT}\n\nExecution policy:\n${this.describeExecutionPolicy(policy)}\n\nCurrent spreadsheet state:\n${context}` },
      { role: 'user', content: instruction },
    ]
    const toolExecutions: AgentToolExecution[] = []
    const plans: AgentPlanPreview[] = []

    try {
      const maxTurns = 12
      for (let turn = 0; turn < maxTurns; turn++) {
        const { content, tool_calls } = await this.callLLM(messages, controller.signal, this.getToolsForPolicy(policy), turn)

        if (!tool_calls || tool_calls.length === 0) {
          return this.buildExecutionResult(content, toolExecutions, plans)
        }

        messages.push({ role: 'assistant', content, tool_calls })

        for (const toolCall of tool_calls) {
          let args: Record<string, unknown>
          try {
            args = JSON.parse(toolCall.function.arguments) as Record<string, unknown>
          } catch {
            messages.push({
              role: 'tool',
              tool_call_id: toolCall.id,
              content: JSON.stringify({ error: `Failed to parse tool arguments: ${toolCall.function.arguments}` }),
            })
            continue
          }

          const result = this.executeTool(toolCall.function.name, args)
          const compactResult = this.compactToolResultForLLM(toolCall.function.name, result)
          this.logToolResultDiagnostics(toolCall.function.name, result, compactResult)
          toolExecutions.push({
            toolName: toolCall.function.name,
            args,
            result,
          })

          const plan = this.extractPlanPreview(result)
          if (plan) {
            plans.push(plan)
          }

          messages.push({
            role: 'tool',
            tool_call_id: toolCall.id,
            content: JSON.stringify(compactResult),
          })
        }
      }

      return this.buildExecutionResult(
        'Operation completed after reaching the maximum number of turns.',
        toolExecutions,
        plans,
      )
    } finally {
      if (this.activeController === controller) {
        this.activeController = null
        this.activeAbortReason = null
        this.activePolicy = null
      }
    }
  }

  async applyPlan(planId: string): Promise<AgentApplyPlanResult> {
    const trimmedPlanId = planId.trim()
    const result = this.executeTool('apply_plan', { plan_id: trimmedPlanId })
    return {
      message: this.summarizeToolResult(result, trimmedPlanId ? `Plan ${trimmedPlanId} applied.` : 'Plan applied.'),
      result,
    }
  }

  private createExecutionPolicy(instruction: string): ExecutionPolicy {
    const text = instruction.toLowerCase()
    const compactText = text.replace(/\s+/g, '')
    const includesAny = (keywords: string[]) => keywords.some((keyword) => compactText.includes(keyword.toLowerCase().replace(/\s+/g, '')))

    const hasFormatIntent = includesAny([
      '格式', '美化', '排版', '边框', '框线', '加边框', '字体', '字号', '颜色', '背景',
      '加粗', '斜体', '下划线', '居中', '对齐', '列宽', '行高', '拉宽', '宽度',
      '展示所有内容', '显示所有内容', '字数较多', '换行', '自动换行',
      '展示', '显示', '小数', '小数点', '保留', '位数', '三位', '两位', '格式化',
      'format', 'border', 'width', 'height', 'rowheight', 'font', 'align', 'wrap',
    ])
    const hasStructureIntent = includesAny([
      '新建工作表', '新增工作表', '创建工作表', '插入行', '新增行', '添加行',
      '插入列', '新增列', '添加列', '删除行', '删行', '删除列', '删列',
      '合并单元格', '合并', '拆分单元格', '取消合并',
      '重命名工作表', '改工作表名', 'rename sheet', 'insert row', 'delete row',
      'insert column', 'delete column', 'new sheet', 'merge cell', 'merge cells',
    ])
    const hasClearIntent = includesAny([
      '清空', '清除', '删除内容', '清除内容', '清除格式', '清空格式',
      'clear', 'remove content', 'delete content',
    ])
    const hasDataIntent = includesAny([
      '写入', '填入', '填写', '填充', '更新', '修改', '改成', '设置为', '计算',
      '同步', '匹配', '重组', '复制', '粘贴', '生成', '汇总', '统计', '公式',
      '金额', '单价', '数量', '总价', '表头', '数据', '内容',
      'write', 'fill', 'update', 'sync', 'match', 'calculate', 'formula',
    ]) || hasClearIntent

    const allowDataTools = hasDataIntent || (!hasFormatIntent && !hasStructureIntent)
    const allowFormatTools = hasFormatIntent || (!hasDataIntent && !hasStructureIntent && !hasClearIntent)
    const allowStructureTools = hasStructureIntent
    const allowClearTools = hasClearIntent

    const intents = [
      allowDataTools ? 'data' : null,
      allowFormatTools ? 'format' : null,
      allowStructureTools ? 'structure' : null,
      allowClearTools ? 'clear' : null,
    ].filter(Boolean).join(' + ')

    return {
      allowDataTools,
      allowFormatTools,
      allowStructureTools,
      allowClearTools,
      intentSummary: intents || 'read-only',
    }
  }

  private describeExecutionPolicy(policy: ExecutionPolicy): string {
    return [
      `Detected intents: ${policy.intentSummary}.`,
      `Data tools: ${policy.allowDataTools ? 'enabled' : 'disabled'}.`,
      `Format tools: ${policy.allowFormatTools ? 'enabled' : 'disabled'}.`,
      `Structure tools: ${policy.allowStructureTools ? 'enabled' : 'disabled unless explicitly requested'}.`,
      `Clear tools: ${policy.allowClearTools ? 'enabled because clear/delete content was explicit' : 'disabled unless clear/delete content is explicit'}.`,
      'If the user request has multiple intents, execute them in safe order: inspect first, data changes second, formatting last.',
    ].join('\n')
  }

  private getToolsForPolicy(policy: ExecutionPolicy): ToolDefinition[] {
    return TOOLS.filter((tool) => this.isToolAllowedByPolicy(tool.function.name, policy))
  }

  private isToolAllowedByPolicy(toolName: string, policy: ExecutionPolicy): boolean {
    const name = toolName as ToolName
    if (READ_TOOL_NAMES.has(name)) return true
    if (name === 'clear_range') return policy.allowClearTools
    if (DATA_TOOL_NAMES.has(name)) return policy.allowDataTools
    if (FORMAT_TOOL_NAMES.has(name)) return policy.allowFormatTools
    if (STRUCTURE_TOOL_NAMES.has(name)) return policy.allowStructureTools
    return true
  }

  private getContext(): string {
    const api = getUniverAPI()
    const workbook = api?.getActiveWorkbook()
    if (!workbook) return 'No workbook'

    const typedWorkbook = workbook as unknown as WorkbookLike
    const snapshot = this.schemaCache.refresh(typedWorkbook)
    const activeSheet = workbook.getActiveSheet()

    let context = `Sheets (${snapshot.sheets.length}): ${snapshot.sheets.map((sheet) => sheet.sheetName).join(', ')}\n`
    context += `Sheet summaries:\n${snapshot.sheets.map((sheet) => {
      const tableSummary = sheet.tables.length > 0
        ? sheet.tables.map((table) => `${table.range} [${table.columns.map((column) => column.header).join(', ')}]`).join('; ')
        : 'no detected tables'
      return `- ${sheet.sheetName}: used=${sheet.usedRange ?? 'empty'}; tables=${tableSummary}`
    }).join('\n')}\n`
    context += `Active sheet: ${snapshot.activeSheetName ?? 'Unknown'}\n`
    context += `Active cell: ${snapshot.activeCell ?? 'Unknown'}\n`
    context += `Active range: ${snapshot.activeRange ?? 'Unknown'}\n`
    context += `Selected ranges: ${snapshot.selectedRanges.length > 0 ? snapshot.selectedRanges.join(', ') : 'none'}\n`

    const activeTables = snapshot.activeSheetName
      ? snapshot.tables.filter((table) => table.sheetName === snapshot.activeSheetName)
      : []
    if (activeTables.length > 0) {
      context += `Detected tables on active sheet:\n${activeTables.map((table) =>
        `- ${table.range}: headers=${table.columns.map((column) => column.header).join(', ')}; keys=${table.primaryKeyCandidates.join(', ') || 'none'}`
      ).join('\n')}\n`
    }

    const selection = activeSheet?.getSelection()
    const selectedRange = selection?.getActiveRangeList()?.[0]
    if (selectedRange) {
      context += `${this.describePreviewRange(selectedRange, 'Selected range preview')}\n`
    }

    if (!activeSheet) {
      return context
    }

    const lastRow = activeSheet.getLastRow()
    const lastColumn = activeSheet.getLastColumn()
    if (lastRow < 0 || lastColumn < 0) {
      return `${context}Sheet is empty.\n`
    }

    const rowCount = Math.min(lastRow + 1, CONTEXT_PREVIEW_MAX_ROWS)
    const columnCount = Math.min(lastColumn + 1, CONTEXT_PREVIEW_MAX_COLUMNS)
    const previewRange = activeSheet.getRange(`A1:${colIndexToLetter(columnCount - 1)}${rowCount}`)
    context += `${this.describePreviewRange(previewRange, 'Top-left preview')}\n`
    return context
  }

  private describePreviewRange(
    range: {
      getA1Notation(withSheet?: boolean): string
      getValues(): unknown[][]
      getFormulas(): string[][]
      getRow(): number
      getColumn(): number
      getHeight(): number
      getWidth(): number
    },
    label: string,
  ): string {
    const height = Math.min(range.getHeight(), CONTEXT_PREVIEW_MAX_ROWS)
    const width = Math.min(range.getWidth(), CONTEXT_PREVIEW_MAX_COLUMNS)
    const values = range.getValues().slice(0, height).map((row) => row.slice(0, width))
    const formulas = range.getFormulas().slice(0, height).map((row) => row.slice(0, width))
    const cellLines = this.serializeGrid(values, formulas, range.getRow(), range.getColumn())

    if (cellLines.length === 0) {
      return `${label}: ${range.getA1Notation()} (empty)`
    }

    return `${label}: ${range.getA1Notation()}\n${cellLines.join('\n')}`
  }

  private serializeGrid(
    values: unknown[][],
    formulas: string[][],
    startRow: number,
    startColumn: number,
  ): string[] {
    const lines: string[] = []

    values.forEach((row, rowOffset) => {
      const cells = row
        .map((value, columnOffset) => {
          const formula = formulas[rowOffset]?.[columnOffset] ?? ''
          const cellRef = `${colIndexToLetter(startColumn + columnOffset)}${startRow + rowOffset + 1}`

          if (formula) {
            return `${cellRef}=${JSON.stringify(formula)} => ${JSON.stringify(value ?? null)}`
          }

          if (value === null || value === '') {
            return ''
          }

          return `${cellRef}=${JSON.stringify(value)}`
        })
        .filter(Boolean)

      if (cells.length > 0) {
        lines.push(cells.join('  '))
      }
    })

    return lines
  }

  private captureSheetSnapshot(sheet: SheetLike, sheetIndex: number): SheetSnapshot {
    const lastRow = sheet.getLastRow()
    const lastColumn = sheet.getLastColumn()
    const sheetName = sheet.getSheetName()
    const sheetId = sheet.getSheetId()
    const mergedRanges = this.readMergedRanges(sheet)
    const mergedBounds = mergedRanges
      .map((range) => this.parseA1Bounds(range))
      .filter((bounds): bounds is { startRow: number; startColumn: number; endRow: number; endColumn: number } => !!bounds)
    const mergedLastRow = mergedBounds.reduce((max, bounds) => Math.max(max, bounds.endRow), -1)
    const mergedLastColumn = mergedBounds.reduce((max, bounds) => Math.max(max, bounds.endColumn), -1)
    const effectiveLastRow = Math.max(lastRow, mergedLastRow)
    const effectiveLastColumn = Math.max(lastColumn, mergedLastColumn)

    if (effectiveLastRow < 0 || effectiveLastColumn < 0) {
      return {
        sheetId,
        sheetName,
        sheetIndex,
        usedRange: null,
        rowCount: 0,
        columnCount: 0,
        values: [],
        formulas: [],
        cellData: [],
        columnWidths: [],
        rowHeights: [],
        mergedRanges,
      }
    }

    const rowCount = effectiveLastRow + 1
    const columnCount = effectiveLastColumn + 1
    const usedRange = rangeToA1(0, 0, effectiveLastRow, effectiveLastColumn)
    const range = sheet.getRange(0, 0, rowCount, columnCount)
    const values = this.cloneJson(range.getValues())
    const formulas = this.cloneJson(range.getFormulas())
    const cellData = this.cloneJson(range.getCellDataGrid())
    const columnWidths = Array.from({ length: columnCount }, (_, columnIndex) => this.readColumnWidth(sheet, columnIndex))
    const rowHeights = Array.from({ length: rowCount }, (_, rowIndex) => this.readRowHeight(sheet, rowIndex))

    return {
      sheetId,
      sheetName,
      sheetIndex,
      usedRange,
      rowCount,
      columnCount,
      values,
      formulas,
      cellData,
      columnWidths,
      rowHeights,
      mergedRanges,
    }
  }

  private readMergedRanges(sheet: SheetLike): string[] {
    try {
      return sheet.getMergedRanges()
        .map((range) => range.getA1Notation())
        .filter((range) => !!range)
    } catch {
      return []
    }
  }

  private readColumnWidth(sheet: SheetLike, columnIndex: number): number {
    try {
      const width = sheet.getColumnWidth(columnIndex)
      return Number.isFinite(width) ? width : 0
    } catch {
      return 0
    }
  }

  private readRowHeight(sheet: SheetLike, rowIndex: number): number {
    try {
      const height = sheet.getRowHeight(rowIndex)
      return Number.isFinite(height) ? height : 0
    } catch {
      return 0
    }
  }

  private readSnapshotCell(
    snapshot: SheetSnapshot | undefined,
    rowIndex: number,
    columnIndex: number,
  ): { value: unknown; formula: string; style: unknown } {
    const cellData = snapshot?.cellData?.[rowIndex]?.[columnIndex] ?? null
    const style = cellData && typeof cellData === 'object' && 's' in cellData ? cellData.s : null
    return {
      value: snapshot?.values?.[rowIndex]?.[columnIndex] ?? null,
      formula: snapshot?.formulas?.[rowIndex]?.[columnIndex] ?? '',
      style,
    }
  }

  private snapshotValuesEqual(left: unknown, right: unknown): boolean {
    if ((left === null || left === undefined || left === '') && (right === null || right === undefined || right === '')) {
      return true
    }

    return JSON.stringify(left ?? null) === JSON.stringify(right ?? null)
  }

  private syncWorkbookSheetsToSnapshot(workbook: WorkbookLike, snapshot: WorkbookSnapshot): void {
    const snapshotSheetNames = new Set(snapshot.sheets.map((sheet) => sheet.sheetName))
    const currentSheets = workbook.getSheets()

    currentSheets.forEach((sheet) => {
      if (!snapshotSheetNames.has(sheet.getSheetName()) && currentSheets.length > 1) {
        try {
          workbook.deleteSheet(sheet)
        } catch {
          try {
            workbook.deleteSheet(sheet.getSheetId())
          } catch {
            // Keep restore best-effort if the facade refuses deletion.
          }
        }
      }
    })

    snapshot.sheets
      .slice()
      .sort((left, right) => left.sheetIndex - right.sheetIndex)
      .forEach((sheetSnapshot) => {
        if (!workbook.getSheetByName(sheetSnapshot.sheetName)) {
          workbook.insertSheet(sheetSnapshot.sheetName)
        }
      })

    workbook.getSheets().forEach((sheet) => this.clearSheetUsedRange(sheet))
  }

  private clearSheetUsedRange(sheet: SheetLike): void {
    this.breakAllMergedRanges(sheet)

    const lastRow = sheet.getLastRow()
    const lastColumn = sheet.getLastColumn()
    if (lastRow < 0 || lastColumn < 0) return

    sheet.getRange(0, 0, lastRow + 1, lastColumn + 1).clear()
  }

  private restoreSheetSnapshot(sheet: SheetLike, snapshot: SheetSnapshot): void {
    this.breakAllMergedRanges(sheet)

    if (snapshot.rowCount > 0 && snapshot.columnCount > 0) {
      const range = sheet.getRange(0, 0, snapshot.rowCount, snapshot.columnCount)
      const cellData = snapshot.cellData.length > 0 ? snapshot.cellData : this.valuesToCellData(snapshot.values, snapshot.formulas)
      range.setValues(cellData)

      snapshot.formulas.forEach((row, rowIndex) => {
        row.forEach((formula, columnIndex) => {
          if (formula) {
            sheet.getRange(rowIndex, columnIndex).setFormula(formula)
          }
        })
      })
    }

    this.restoreColumnWidths(sheet, snapshot.columnWidths)
    this.restoreRowHeights(sheet, snapshot.rowHeights)
    this.restoreMergedRanges(sheet, snapshot.mergedRanges)
  }

  private breakAllMergedRanges(sheet: SheetLike): void {
    try {
      sheet.getMergedRanges().forEach((range) => range.breakApart())
    } catch {
      // Older or partial facade implementations may not expose merge inspection.
    }
  }

  private restoreColumnWidths(sheet: SheetLike, columnWidths: number[]): void {
    columnWidths.forEach((width, columnIndex) => {
      if (Number.isFinite(width) && width > 0) {
        sheet.setColumnWidth(columnIndex, width)
      }
    })
  }

  private restoreRowHeights(sheet: SheetLike, rowHeights: number[]): void {
    rowHeights.forEach((height, rowIndex) => {
      if (Number.isFinite(height) && height > 0) {
        sheet.setRowHeight(rowIndex, height)
      }
    })
  }

  private restoreMergedRanges(sheet: SheetLike, mergedRanges: string[]): void {
    mergedRanges.forEach((rangeA1) => {
      try {
        sheet.getRange(rangeA1).merge({ isForceMerge: true })
      } catch {
        // Keep snapshot restore best-effort; values/styles/size restoration should still succeed.
      }
    })
  }

  private valuesToCellData(values: unknown[][], formulas: string[][]): Array<Array<Record<string, unknown> | null>> {
    return values.map((row, rowIndex) =>
      row.map((value, columnIndex) => {
        const formula = formulas[rowIndex]?.[columnIndex]
        if (formula) return { f: formula }
        if (value === null || value === undefined || value === '') return null
        return { v: value }
      }),
    )
  }

  private cloneJson<T>(value: T): T {
    return JSON.parse(JSON.stringify(value)) as T
  }

  private async callLLM(
    messages: Message[],
    signal: AbortSignal,
    tools: ToolDefinition[],
    turn: number,
  ): Promise<{ content: string | null; tool_calls?: ToolCall[] }> {
    const url = `${API_BASE}/api/llm/chat-completions`
    const timeoutId = window.setTimeout(() => {
      this.activeAbortReason = `LLM request timed out after ${UniverAgent.REQUEST_TIMEOUT_MS / 1000} seconds`
      this.activeController?.abort()
    }, UniverAgent.REQUEST_TIMEOUT_MS)

    try {
      const headers: Record<string, string> = { 'Content-Type': 'application/json' }

      return await this.sendLLMRequestWithFallback(url, headers, messages, tools, signal, turn)
    } catch (error) {
      if (signal.aborted) {
        throw new Error(this.activeAbortReason ?? 'Request canceled', { cause: error })
      }
      throw error
    } finally {
      window.clearTimeout(timeoutId)
    }

  }

  private buildLLMPayload(
    messages: Message[],
    tools: ToolDefinition[],
    mode: LLMPayloadMode,
  ): Record<string, unknown> {
    const compatibilityMode = mode !== 'standard'
    const payload: Record<string, unknown> = {
      model: this.config.model,
      messages: this.normalizeMessagesForLLM(messages, compatibilityMode),
      stream: false,
    }

    const normalizedTools = this.normalizeToolsForLLM(
      mode === 'minimal' ? this.filterMinimalCompatibleTools(tools) : tools,
      compatibilityMode,
    )
    if (normalizedTools.length > 0) {
      payload.tools = normalizedTools
      payload.tool_choice = 'auto'
    }

    return payload
  }

  private async sendLLMRequestWithFallback(
    url: string,
    headers: Record<string, string>,
    messages: Message[],
    tools: ToolDefinition[],
    signal: AbortSignal,
    turn: number,
  ): Promise<{ content: string | null; tool_calls?: ToolCall[] }> {
    const modes: LLMPayloadMode[] = ['standard', 'compact', 'minimal']
    let lastError: unknown = null

    for (const mode of modes) {
      if (mode !== 'standard' && lastError && !this.shouldRetryInCompatibilityMode(lastError)) {
        break
      }

      const payload = this.buildLLMPayload(messages, tools, mode)
      try {
        return await this.sendLLMRequest(url, headers, payload, signal, { turn, mode })
      } catch (error) {
        lastError = error
        this.logLLMRequestError(error, turn, mode)
      }
    }

    throw lastError
  }

  private async sendLLMRequest(
    url: string,
    headers: Record<string, string>,
    payload: Record<string, unknown>,
    signal: AbortSignal,
    diagnostics: { turn: number; mode: LLMPayloadMode },
  ): Promise<{ content: string | null; tool_calls?: ToolCall[] }> {
    const body = JSON.stringify(payload)
    this.logLLMPayloadDiagnostics(payload, body, diagnostics.turn, diagnostics.mode)
    const proxyBody = JSON.stringify({
      baseURL: this.config.baseURL,
      apiKey: this.config.apiKey,
      payload,
    })
    const response = await fetch(url, {
      method: 'POST',
      headers,
      credentials: 'include',
      body: proxyBody,
      signal,
    })

    if (!response.ok) {
      const text = await response.text()
      throw new Error(`LLM API ${response.status}: ${text.slice(0, 400)}`)
    }

    const data = await response.json()
    const message = data.choices?.[0]?.message
    if (!message) throw new Error('Unexpected LLM response shape')

    return {
      content: this.extractLLMMessageContent(message.content),
      tool_calls: Array.isArray(message.tool_calls) ? message.tool_calls : undefined,
    }
  }

  private shouldRetryInCompatibilityMode(error: unknown): boolean {
    const message = error instanceof Error ? error.message : String(error)
    const normalized = message.toLowerCase()
    return normalized.includes('upstream_empty_output')
      || normalized.includes('invalid parameters')
      || normalized.includes('invalid schema')
      || normalized.includes('tool_choice')
  }

  private normalizeMessagesForLLM(messages: Message[], compatibilityMode = false): LLMRequestMessage[] {
    return messages.map((message) => {
      if (message.role === 'assistant') {
        return {
          role: 'assistant',
          content: this.normalizeLLMMessageContent(typeof message.content === 'string' ? message.content : '', compatibilityMode),
          ...(Array.isArray(message.tool_calls) && message.tool_calls.length > 0 ? { tool_calls: message.tool_calls } : {}),
        }
      }

      if (message.role === 'tool') {
        return {
          role: 'tool',
          tool_call_id: message.tool_call_id,
          content: this.normalizeLLMMessageContent(message.content, compatibilityMode),
        }
      }

      return {
        role: message.role,
        content: this.normalizeLLMMessageContent(message.content, compatibilityMode),
      }
    })
  }

  private normalizeLLMMessageContent(content: string, compatibilityMode: boolean): string {
    const normalized = String(content ?? '')
    if (!compatibilityMode) {
      return normalized
    }

    const maxLength = 12_000
    if (normalized.length <= maxLength) {
      return normalized
    }

    return `${normalized.slice(0, maxLength)}\n...[truncated for compatibility]`
  }

  private normalizeToolsForLLM(tools: ToolDefinition[], compatibilityMode: boolean): LLMToolDefinition[] {
    return tools.map((tool) => ({
      type: 'function',
      function: {
        name: tool.function.name,
        description: compatibilityMode
          ? this.compactToolDescription(tool.function.description)
          : tool.function.description,
        parameters: this.normalizeToolSchema(tool.function.parameters, compatibilityMode),
      },
    }))
  }

  private filterMinimalCompatibleTools(tools: ToolDefinition[]): ToolDefinition[] {
    const filtered = tools.filter((tool) => MINIMAL_COMPATIBLE_TOOL_NAMES.has(tool.function.name as ToolName))
    return filtered.length > 0 ? filtered : tools
  }

  private logLLMPayloadDiagnostics(
    payload: Record<string, unknown>,
    body: string,
    turn: number,
    mode: LLMPayloadMode,
  ): void {
    const messages = Array.isArray(payload.messages) ? payload.messages : []
    const tools = Array.isArray(payload.tools) ? payload.tools as LLMToolDefinition[] : []
    const messageChars = messages.reduce((total, message) => {
      if (!this.isPlainObject(message)) return total
      return total + (typeof message.content === 'string' ? message.content.length : 0)
    }, 0)

    console.info('[univer-agent][llm-payload]', {
      turn,
      mode,
      bytes: body.length,
      messageCount: messages.length,
      messageChars,
      toolCount: tools.length,
      toolNames: tools.map((tool) => tool.function.name),
    })
  }

  private logLLMRequestError(error: unknown, turn: number, mode: LLMPayloadMode): void {
    const message = error instanceof Error ? error.message : String(error)
    console.warn('[univer-agent][llm-error]', {
      turn,
      mode,
      message: message.slice(0, 600),
    })
  }

  private logToolResultDiagnostics(toolName: string, result: unknown, compactResult: unknown): void {
    console.info('[univer-agent][tool-result]', {
      toolName,
      rawBytes: this.safeJsonLength(result),
      compactBytes: this.safeJsonLength(compactResult),
    })
  }

  private safeJsonLength(value: unknown): number {
    try {
      return JSON.stringify(value).length
    } catch {
      return -1
    }
  }

  private compactToolDescription(description: string): string {
    const normalized = description.replace(/\s+/g, ' ').trim()
    return normalized.length > 160 ? `${normalized.slice(0, 157)}...` : normalized
  }

  private normalizeToolSchema(schema: unknown, compatibilityMode: boolean): Record<string, unknown> {
    return this.sanitizeSchemaNode(schema, compatibilityMode) as Record<string, unknown>
  }

  private sanitizeSchemaNode(schema: unknown, compatibilityMode: boolean): Record<string, unknown> {
    if (!this.isPlainObject(schema)) {
      return { type: 'string' }
    }

    const explicitType = typeof schema.type === 'string' ? schema.type : ''
    const inferredType = explicitType
      || (this.isPlainObject(schema.properties) || 'required' in schema || 'additionalProperties' in schema ? 'object' : '')
      || ('items' in schema ? 'array' : '')
      || (Array.isArray(schema.enum) && schema.enum.length > 0 ? this.inferEnumSchemaType(schema.enum) : '')
      || 'string'

    const result: Record<string, unknown> = { type: inferredType }

    if (typeof schema.description === 'string' && schema.description.trim()) {
      result.description = compatibilityMode
        ? this.compactToolDescription(schema.description)
        : schema.description
    }

    if (Array.isArray(schema.enum) && schema.enum.length > 0 && inferredType !== 'object' && inferredType !== 'array') {
      result.enum = schema.enum
    }

    if (inferredType === 'object') {
      const properties = this.isPlainObject(schema.properties) ? schema.properties : {}
      const nextProperties = Object.fromEntries(
        Object.entries(properties).map(([key, value]) => [key, this.sanitizeSchemaNode(value, compatibilityMode)]),
      )

      result.properties = nextProperties
      result.required = Array.isArray(schema.required)
        ? schema.required.filter((item): item is string => typeof item === 'string')
        : []

      if (schema.additionalProperties === true) {
        result.additionalProperties = { type: 'string' }
      } else if (this.isPlainObject(schema.additionalProperties)) {
        result.additionalProperties = this.sanitizeSchemaNode(schema.additionalProperties, compatibilityMode)
      } else if (compatibilityMode) {
        result.additionalProperties = false
      }
    }

    if (inferredType === 'array') {
      result.items = this.sanitizeSchemaNode(schema.items, compatibilityMode)
    }

    return result
  }

  private inferEnumSchemaType(values: unknown[]): string {
    const first = values.find((value) => value !== null && value !== undefined)
    if (typeof first === 'number') return Number.isInteger(first) ? 'integer' : 'number'
    if (typeof first === 'boolean') return 'boolean'
    return 'string'
  }

  private isPlainObject(value: unknown): value is Record<string, unknown> {
    return !!value && typeof value === 'object' && !Array.isArray(value)
  }

  private compactToolResultForLLM(toolName: string, result: unknown): unknown {
    const compacted = this.compactValueForLLM(result, 0)
    const serialized = JSON.stringify(compacted)

    if (serialized.length <= UniverAgent.MAX_TOOL_RESULT_JSON_LENGTH) {
      return compacted
    }

    return {
      tool: toolName,
      notice: 'Tool result truncated for LLM compatibility.',
      summary: this.summarizeToolResult(result, `${toolName} completed.`),
      compacted: this.compactValueForLLM(compacted, 1),
    }
  }

  private compactValueForLLM(value: unknown, depth: number): unknown {
    if (value === null || value === undefined) {
      return value ?? null
    }

    if (typeof value === 'string') {
      if (value.length <= UniverAgent.MAX_TOOL_RESULT_STRING_LENGTH) {
        return value
      }

      return `${value.slice(0, UniverAgent.MAX_TOOL_RESULT_STRING_LENGTH)}...[truncated]`
    }

    if (typeof value === 'number' || typeof value === 'boolean') {
      return value
    }

    if (Array.isArray(value)) {
      const limit = depth <= 1 ? UniverAgent.MAX_TOOL_RESULT_ARRAY_ITEMS : 8
      const sliced = value.slice(0, limit).map((item) => this.compactValueForLLM(item, depth + 1))
      if (value.length > limit) {
        sliced.push({
          __truncated__: true,
          omitted_items: value.length - limit,
        })
      }
      return sliced
    }

    if (typeof value === 'object') {
      const entries = Object.entries(value as Record<string, unknown>)
      const limit = depth === 0 ? 40 : 20
      const compacted = Object.fromEntries(
        entries.slice(0, limit).map(([key, item]) => [key, this.compactValueForLLM(item, depth + 1)]),
      )

      if (entries.length > limit) {
        compacted.__truncated_keys__ = entries.length - limit
      }

      return compacted
    }

    return String(value)
  }

  private compactMatrixForLLM(matrix: unknown[][]): unknown[][] {
    return matrix
      .slice(0, CONTEXT_PREVIEW_MAX_ROWS)
      .map((row) => row.slice(0, CONTEXT_PREVIEW_MAX_COLUMNS))
  }

  private extractLLMMessageContent(content: unknown): string | null {
    if (typeof content === 'string') {
      return content
    }

    if (Array.isArray(content)) {
      const text = content
        .map((item) => {
          if (typeof item === 'string') return item
          if (!item || typeof item !== 'object') return ''

          const textValue = 'text' in item && typeof item.text === 'string'
            ? item.text
            : ('content' in item && typeof item.content === 'string' ? item.content : '')

          return textValue
        })
        .filter(Boolean)
        .join('\n')
        .trim()

      return text || null
    }

    return null
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

  private applyRangeFormatting(
    range: unknown,
    args: Record<string, unknown>,
  ): void {
    const targetRange = range as {
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
      getCellDataGrid(): Array<Array<Record<string, unknown> | null>>
      getValues(): unknown[][]
      getFormulas(): string[][]
      setValues(values: Array<Array<Record<string, unknown> | null>>): void
    }

    if (args.bold !== undefined) targetRange.setFontWeight(args.bold ? 'bold' : 'normal')
    if (args.italic !== undefined) targetRange.setFontStyle(args.italic ? 'italic' : 'normal')
    if (typeof args.font_family === 'string' && args.font_family.trim() && targetRange.setFontFamily) {
      targetRange.setFontFamily(args.font_family.trim())
    }
    if (args.font_size !== undefined) targetRange.setFontSize(args.font_size as number)
    if (args.font_color !== undefined) targetRange.setFontColor(args.font_color as string)
    if (args.background !== undefined) targetRange.setBackground(args.background as string)
    if (args.underline !== undefined) targetRange.setFontLine(args.underline ? 'underline' : 'none')

    if (args.horizontal_alignment !== undefined) {
      const alignmentMap: Record<string, 'left' | 'center' | 'normal'> = {
        left: 'left',
        center: 'center',
        right: 'normal',
      }
      targetRange.setHorizontalAlignment(alignmentMap[String(args.horizontal_alignment)] ?? 'left')
    }

    if (typeof args.vertical_alignment === 'string' && targetRange.setVerticalAlignment) {
      const alignmentMap: Record<string, 'top' | 'middle' | 'bottom'> = {
        top: 'top',
        middle: 'middle',
        bottom: 'bottom',
      }
      targetRange.setVerticalAlignment(alignmentMap[String(args.vertical_alignment)] ?? 'middle')
    }

    if (typeof args.text_rotation === 'number' && Number.isFinite(args.text_rotation) && targetRange.setTextRotation) {
      targetRange.setTextRotation(Math.max(-90, Math.min(90, Math.round(args.text_rotation))))
    }

    if (args.wrap_strategy !== undefined) {
      const univerEnum = getUniverAPI()?.Enum as Record<string, Record<string, unknown>> | undefined
      const wrapKey = String(args.wrap_strategy).trim().toUpperCase()
      const wrapStrategy = univerEnum?.WrapStrategy?.[wrapKey] ?? univerEnum?.WrapStrategy?.WRAP
      if (wrapStrategy !== undefined) {
        targetRange.setWrapStrategy(wrapStrategy)
      }
    }

    if (typeof args.number_format === 'string' && args.number_format.trim()) {
      const pattern = args.number_format.trim()
      const cellData = targetRange.getCellDataGrid()
      const values = targetRange.getValues()
      const formulas = targetRange.getFormulas()
      const nextCellData = values.map((row, rowIndex) =>
        row.map((value, columnIndex) => {
          const existingCell = cellData[rowIndex]?.[columnIndex]
          const formula = formulas[rowIndex]?.[columnIndex]
          const nextCell: Record<string, unknown> = existingCell && typeof existingCell === 'object'
            ? { ...existingCell }
            : formula
              ? { f: formula }
              : value === null || value === undefined || value === ''
                ? {}
                : { v: value }
          const existingStyle = nextCell.s && typeof nextCell.s === 'object' && !Array.isArray(nextCell.s)
            ? { ...(nextCell.s as Record<string, unknown>) }
            : {}
          nextCell.s = {
            ...existingStyle,
            n: { pattern },
          }
          return nextCell
        }),
      )
      targetRange.setValues(nextCellData)
    }

    if (args.border_type !== undefined) {
      const univerEnum = getUniverAPI()?.Enum as Record<string, Record<string, unknown>> | undefined
      const borderTypeKey = String(args.border_type).trim().toUpperCase()
      const borderStyleKey = String(args.border_style ?? 'thin').trim().toUpperCase()
      const borderType = univerEnum?.BorderType?.[borderTypeKey] ?? univerEnum?.BorderType?.ALL
      const borderStyle = univerEnum?.BorderStyleTypes?.[borderStyleKey] ?? univerEnum?.BorderStyleTypes?.THIN
      if (borderType !== undefined && borderStyle !== undefined) {
        targetRange.setBorder(borderType, borderStyle, typeof args.border_color === 'string' ? args.border_color : '#000000')
      }
    }
  }

  private resolveRangeA1(sheet: unknown, args: Record<string, unknown>): string | { error: string } {
    const targetSheet = sheet as {
      getLastRow(): number
      getLastColumn(): number
    }

    if (typeof args.range === 'string' && args.range.trim()) {
      return args.range.trim()
    }

    const activeRange = getUniverAPI()?.getActiveWorkbook()?.getActiveRange?.()
    if (activeRange) {
      return activeRange.getA1Notation()
    }

    const lastRow = targetSheet.getLastRow()
    const lastColumn = targetSheet.getLastColumn()
    if (lastRow >= 0 && lastColumn >= 0) {
      return rangeToA1(0, 0, lastRow, lastColumn)
    }

    return { error: 'range is required because the sheet is empty and no active range is available.' }
  }

  private resolveRowSpan(sheet: unknown, args: Record<string, unknown>): { rowIndex: number; count: number; range: string | null } | { error: string } {
    if (typeof args.row_index === 'number') {
      const rowIndex = Math.floor(args.row_index)
      const count = typeof args.count === 'number' ? Math.max(1, Math.floor(args.count)) : 1
      return { rowIndex, count, range: null }
    }

    const resolvedRange = this.resolveRangeA1(sheet, args)
    if (typeof resolvedRange !== 'string') {
      return resolvedRange
    }

    const targetSheet = sheet as {
      getRange(a1: string): {
        getRange(): { startRow: number; startColumn: number; endRow: number; endColumn: number } | undefined
      }
    }
    const bounds = targetSheet.getRange(resolvedRange).getRange() ?? this.parseA1Bounds(resolvedRange)
    if (!bounds) {
      return { error: `Invalid range: ${resolvedRange}` }
    }

    return {
      rowIndex: bounds.startRow,
      count: bounds.endRow - bounds.startRow + 1,
      range: resolvedRange,
    }
  }

  private parseA1Bounds(rangeA1: string): { startRow: number; startColumn: number; endRow: number; endColumn: number } | null {
    const cleaned = rangeA1
      .trim()
      .replace(/^'[^']+'!/, '')
      .replace(/^[^!]+!/, '')
      .replace(/\$/g, '')
      .toUpperCase()

    const [startRef, endRef = startRef] = cleaned.split(':')
    const start = /^([A-Z]+)(\d+)$/.exec(startRef)
    const end = /^([A-Z]+)(\d+)$/.exec(endRef)
    if (!start || !end) return null

    const startRow = Number(start[2]) - 1
    const startColumn = tableColumnLetterToIndex(start[1])
    const endRow = Number(end[2]) - 1
    const endColumn = tableColumnLetterToIndex(end[1])

    if (startRow < 0 || startColumn < 0 || endRow < 0 || endColumn < 0) {
      return null
    }

    return {
      startRow: Math.min(startRow, endRow),
      startColumn: Math.min(startColumn, endColumn),
      endRow: Math.max(startRow, endRow),
      endColumn: Math.max(startColumn, endColumn),
    }
  }

  private autoResizeColumnsForRange(sheet: unknown, args: Record<string, unknown>): { columns: Array<{ column_index: number; width: number }>; range: string } | { error: string } {
    const targetSheet = sheet as {
      getRange(a1: string): {
        getRange(): { startRow: number; startColumn: number; endRow: number; endColumn: number } | undefined
        getValues(): unknown[][]
      }
      setColumnWidth(columnIndex: number, width: number): void
    }

    const resolvedRange = this.resolveRangeA1(sheet, args)
    if (typeof resolvedRange !== 'string') {
      return resolvedRange
    }

    const minWidth = typeof args.min_width === 'number' ? Math.max(40, args.min_width) : 80
    const maxWidth = typeof args.max_width === 'number' ? Math.max(minWidth, args.max_width) : 420
    const padding = typeof args.padding === 'number' ? Math.max(0, args.padding) : 28

    const range = targetSheet.getRange(resolvedRange)
    const bounds = range.getRange() ?? this.parseA1Bounds(resolvedRange)
    if (!bounds) {
      return { error: `Invalid range: ${resolvedRange}` }
    }

    const values = range.getValues()
    const columns: Array<{ column_index: number; width: number }> = []

    for (let columnOffset = 0; columnOffset <= bounds.endColumn - bounds.startColumn; columnOffset++) {
      let longest = 0
      for (const row of values) {
        const raw = row?.[columnOffset]
        const text = raw === null || raw === undefined ? '' : String(raw)
        const weightedLength = [...text].reduce((sum, char) => sum + (char.charCodeAt(0) > 255 ? 1.8 : 1), 0)
        longest = Math.max(longest, weightedLength)
      }

      const width = Math.max(minWidth, Math.min(maxWidth, Math.round(longest * 10 + padding)))
      const columnIndex = bounds.startColumn + columnOffset
      targetSheet.setColumnWidth(columnIndex, width)
      columns.push({ column_index: columnIndex, width })
    }

    return { columns, range: resolvedRange }
  }

  private autoResizeRowsForRange(sheet: unknown, args: Record<string, unknown>): { rows: Array<{ row_index: number; height: number }>; range: string } | { error: string } {
    const targetSheet = sheet as {
      getRange(a1: string): {
        getRange(): { startRow: number; startColumn: number; endRow: number; endColumn: number } | undefined
        getValues(): unknown[][]
      }
      getColumnWidth(columnIndex: number): number
      setRowHeight(rowIndex: number, height: number): void
    }

    const resolvedRange = this.resolveRangeA1(sheet, args)
    if (typeof resolvedRange !== 'string') {
      return resolvedRange
    }

    const minHeight = typeof args.min_height === 'number' ? Math.max(8, args.min_height) : 22
    const maxHeight = typeof args.max_height === 'number' ? Math.max(minHeight, args.max_height) : 180
    const baseLineHeight = typeof args.base_line_height === 'number' ? Math.max(10, args.base_line_height) : 20
    const range = targetSheet.getRange(resolvedRange)
    const bounds = range.getRange() ?? this.parseA1Bounds(resolvedRange)
    if (!bounds) {
      return { error: `Invalid range: ${resolvedRange}` }
    }

    const values = range.getValues()
    const rows: Array<{ row_index: number; height: number }> = []

    values.forEach((row, rowOffset) => {
      let maxLines = 1

      row.forEach((raw, columnOffset) => {
        const text = raw === null || raw === undefined ? '' : String(raw)
        const explicitLines = text.split(/\r\n|\r|\n/).length
        const weightedLength = [...text].reduce((sum, char) => sum + (char.charCodeAt(0) > 255 ? 1.8 : 1), 0)
        const columnIndex = bounds.startColumn + columnOffset
        const columnWidth = safeFiniteNumber(targetSheet.getColumnWidth(columnIndex), 80)
        const charsPerLine = Math.max(4, Math.floor((columnWidth - 12) / 8))
        const wrappedLines = Math.max(1, Math.ceil(weightedLength / charsPerLine))
        maxLines = Math.max(maxLines, explicitLines, wrappedLines)
      })

      const height = Math.max(minHeight, Math.min(maxHeight, Math.round(maxLines * baseLineHeight + 6)))
      const rowIndex = bounds.startRow + rowOffset
      targetSheet.setRowHeight(rowIndex, height)
      rows.push({ row_index: rowIndex, height })
    })

    return { rows, range: resolvedRange }
  }

  private getTableSignature(table: TableDescriptor): string {
    return JSON.stringify({
      sheet: table.sheetName,
      range: table.range,
      headers: table.columns.map((column) => column.header),
      rowCount: table.rowCount,
    })
  }

  private buildExecutionResult(
    content: string | null,
    tools: AgentToolExecution[],
    plans: AgentPlanPreview[],
  ): AgentExecutionResult {
    return {
      message: this.summarizeExecution(content, tools, plans),
      finalMessage: content?.trim() ? content.trim() : null,
      tools,
      plans,
    }
  }

  private summarizeExecution(
    content: string | null,
    tools: AgentToolExecution[],
    plans: AgentPlanPreview[],
  ): string {
    const trimmedContent = content?.trim()
    if (trimmedContent) {
      return trimmedContent
    }

    if (plans.length > 0) {
      return `Generated ${plans.length} preview plan(s). Review the latest preview before applying it.`
    }

    const lastTool = tools.at(-1)
    if (lastTool) {
      return this.summarizeToolResult(lastTool.result, `${lastTool.toolName} completed.`)
    }

    return 'Operation completed.'
  }

  private summarizeToolResult(result: unknown, fallbackMessage: string): string {
    if (!result || typeof result !== 'object' || Array.isArray(result)) {
      return fallbackMessage
    }

    if ('error' in result && typeof result.error === 'string' && result.error.trim()) {
      return result.error
    }

    if ('message' in result && typeof result.message === 'string' && result.message.trim()) {
      return result.message
    }

    if ('summary' in result && typeof result.summary === 'string' && result.summary.trim()) {
      return result.summary
    }

    if ('success' in result && result.success === true) {
      return fallbackMessage
    }

    return fallbackMessage
  }

  private extractPlanPreview(result: unknown): AgentPlanPreview | null {
    if (!result || typeof result !== 'object' || Array.isArray(result)) {
      return null
    }

    const planId = 'plan_id' in result && typeof result.plan_id === 'string' ? result.plan_id.trim() : ''
    const type = 'type' in result && typeof result.type === 'string' ? result.type.trim() : ''
    const summary = 'summary' in result && typeof result.summary === 'string' ? result.summary.trim() : ''

    if (!planId || !type || !summary) {
      return null
    }

    return {
      planId,
      type,
      summary,
      result: result as Record<string, unknown>,
    }
  }

  private executeTool(name: string, args: Record<string, unknown>): unknown {
    if (this.activePolicy && !this.isToolAllowedByPolicy(name, this.activePolicy)) {
      return {
        error: `Tool ${name} is disabled for this request by the execution policy (${this.activePolicy.intentSummary}).`,
      }
    }

    const api = getUniverAPI()
    if (!api) return { error: 'Univer API is not initialized' }

    const workbook = api.getActiveWorkbook()
    if (!workbook) return { error: 'No active workbook' }

    const typedWorkbook = workbook as unknown as WorkbookLike
    const snapshot = this.schemaCache.getSnapshot(typedWorkbook)

    const getSheet = () => {
      if (typeof args.sheet_name === 'string' && args.sheet_name.trim()) {
        const resolvedSheetName = resolveSnapshotSheetName(snapshot, args.sheet_name) ?? args.sheet_name.trim()
        const sheet = findSheetByName(typedWorkbook, resolvedSheetName)
        if (!sheet) {
          throw new Error(`Sheet not found: ${args.sheet_name}`)
        }
        return sheet
      }

      const activeSheet = workbook.getActiveSheet()
      if (!activeSheet) throw new Error('No active sheet')
      return activeSheet
    }

    try {
      switch (name) {
        case 'get_spreadsheet_info': {
          const activeSheet = workbook.getActiveSheet()
          const sheets = workbook.getSheets()
          const selection = activeSheet?.getSelection()
          const selectedRanges = selection?.getActiveRangeList() ?? []
          const lastRow = activeSheet?.getLastRow() ?? -1
          const lastColumn = activeSheet?.getLastColumn() ?? -1
          const rowCount = Math.min(lastRow + 1, CONTEXT_PREVIEW_MAX_ROWS)
          const columnCount = Math.min(lastColumn + 1, CONTEXT_PREVIEW_MAX_COLUMNS)

          let preview: unknown[][] = []
          let previewFormulas: string[][] = []
          if (activeSheet && rowCount > 0 && columnCount > 0) {
            const range = `A1:${colIndexToLetter(columnCount - 1)}${rowCount}`
            const topLeftRange = activeSheet.getRange(range)
            preview = topLeftRange.getValues()
            previewFormulas = topLeftRange.getFormulas()
          }

          const selectionPreview = selectedRanges[0]
            ? {
                range: selectedRanges[0].getA1Notation(),
                values: this.compactMatrixForLLM(selectedRanges[0].getValues()),
                formulas: this.compactMatrixForLLM(selectedRanges[0].getFormulas()),
              }
            : null

          return {
            active_sheet: snapshot.activeSheetName,
            active_cell: snapshot.activeCell,
            active_range: snapshot.activeRange,
            selected_ranges: snapshot.selectedRanges,
            sheets: sheets.map((sheet) => sheet.getSheetName()),
            sheet_summaries: snapshot.sheets.map((sheet) => ({
              sheet_name: sheet.sheetName,
              used_range: sheet.usedRange,
              detected_tables: sheet.tables.map((table) => summarizeTable(table)),
            })),
            tables: snapshot.tables.map((table) => summarizeTable(table)),
            last_row: lastRow,
            last_column: lastColumn,
            used_range: lastRow >= 0 && lastColumn >= 0
              ? `A1:${colIndexToLetter(lastColumn)}${lastRow + 1}`
              : null,
            data_preview: preview,
            data_preview_formulas: previewFormulas,
            selection_preview: selectionPreview,
          }
        }

        case 'list_tables': {
          const sheetName = resolveSnapshotSheetName(snapshot, args.sheet_name) ?? undefined
          const tables = sheetName
            ? snapshot.tables.filter((table) => table.sheetName === sheetName)
            : snapshot.tables
          return {
            sheet_name: sheetName ?? snapshot.activeSheetName,
            tables: tables.map((table) => summarizeTable(table)),
          }
        }

        case 'describe_table': {
          const table = resolveTable(snapshot, args)
          if (!table) {
            return { error: 'No semantic table was detected for the requested sheet or range.' }
          }

          return {
            table: summarizeTable(table),
            total_rows: table.dataRows.length,
            rows: table.dataRows.slice(0, UniverAgent.MAX_TOOL_RESULT_ARRAY_ITEMS).map((rowValues, index) => ({
              row_number: table.headerRowIndex + 2 + index,
              record: this.rowArrayToRecord(table, rowValues),
            })),
            truncated_rows: Math.max(0, table.dataRows.length - UniverAgent.MAX_TOOL_RESULT_ARRAY_ITEMS),
          }
        }

        case 'find_rows_by_conditions': {
          const table = resolveTable(snapshot, args)
          if (!table) {
            return { error: 'No semantic table was detected for the requested sheet or range.' }
          }

          const rowConditions = this.readRowConditions(args.row_conditions)
          const limit = typeof args.limit === 'number' && args.limit > 0 ? args.limit : 5
          const result = findRowsByConditions(table, rowConditions, limit)
          return {
            table: summarizeTable(table),
            resolved_conditions: result.resolved_conditions,
            unresolved_conditions: result.unresolved_conditions,
            matches: result.matches.map((match) => ({
              row_number: match.rowNumber,
              score: match.score,
              record: match.record,
            })),
          }
        }

        case 'locate_target_cell': {
          const result = locateTargetCell(snapshot, args)
          if (result.error) {
            return {
              error: result.error,
              resolved_conditions: result.resolved_conditions,
              unresolved_conditions: result.unresolved_conditions,
              candidates: result.ambiguity?.map((match) => ({
                row_number: match.rowNumber,
                score: match.score,
                record: match.record,
              })),
            }
          }

          return result.location
        }

        case 'preview_set_table_cell_value': {
          const result = locateTargetCell(snapshot, args)
          if (!result.location) {
            return {
              error: result.error ?? 'Unable to locate the target cell.',
              resolved_conditions: result.resolved_conditions,
              unresolved_conditions: result.unresolved_conditions,
              candidates: result.ambiguity?.map((match) => ({
                row_number: match.rowNumber,
                score: match.score,
                record: match.record,
              })),
            }
          }

          const plan = this.planStore.save({
            id: createPlanId(),
            type: 'set_table_cell_value',
            summary: `Update ${result.location.sheet_name}!${result.location.target_cell} from ${JSON.stringify(result.location.current_value)} to ${JSON.stringify(args.value ?? null)}`,
            sheetName: result.location.sheet_name,
            tableRange: result.location.table_range,
            targetCell: result.location.target_cell,
            currentValue: result.location.current_value,
            currentFormula: result.location.current_formula,
            nextValue: args.value ?? null,
          })

          return {
            plan_id: plan.id,
            type: plan.type,
            summary: plan.summary,
            location: result.location,
          }
        }

        case 'set_table_cell_value': {
          const result = locateTargetCell(snapshot, args)
          if (!result.location) {
            return {
              error: result.error ?? 'Unable to locate the target cell.',
              resolved_conditions: result.resolved_conditions,
              unresolved_conditions: result.unresolved_conditions,
              candidates: result.ambiguity?.map((match) => ({
                row_number: match.rowNumber,
                score: match.score,
                record: match.record,
              })),
            }
          }

          const writeResult = this.executor.setTableCellValue(typedWorkbook, result.location, args.value)
          if ('error' in writeResult) {
            return writeResult
          }

          return {
            ...writeResult,
            location: result.location,
          }
        }

        case 'preview_sync_table_to_table': {
          const sourceTable = resolveTableByPrefix(snapshot, args, 'source')
          const targetTable = resolveTableByPrefix(snapshot, args, 'target')
          if (!sourceTable) {
            return { error: 'No source semantic table was detected for the requested source sheet or range.' }
          }
          if (!targetTable) {
            return { error: 'No target semantic table was detected for the requested target sheet or range.' }
          }

          const preview = this.executor.syncTableToTable(typedWorkbook, sourceTable, targetTable, {
            ...args,
            __preview_only: true,
          })
          if (!preview || typeof preview !== 'object' || !('success' in preview) || !preview.success) {
            return preview
          }

          const plan = this.planStore.save({
            id: createPlanId(),
            type: 'sync_table_to_table',
            summary: `Sync ${sourceTable.sheetName}!${sourceTable.range} to ${targetTable.sheetName}!${targetTable.range}`,
            sourceSheetName: sourceTable.sheetName,
            sourceTableRange: sourceTable.range,
            targetSheetName: targetTable.sheetName,
            targetTableRange: targetTable.range,
            args,
            sourceSignature: this.getTableSignature(sourceTable),
            targetSignature: this.getTableSignature(targetTable),
          })

          return {
            plan_id: plan.id,
            type: plan.type,
            summary: plan.summary,
            preview,
          }
        }

        case 'sync_table_to_table': {
          const sourceTable = resolveTableByPrefix(snapshot, args, 'source')
          const targetTable = resolveTableByPrefix(snapshot, args, 'target')
          if (!sourceTable) {
            return { error: 'No source semantic table was detected for the requested source sheet or range.' }
          }
          if (!targetTable) {
            return { error: 'No target semantic table was detected for the requested target sheet or range.' }
          }

          return this.executor.syncTableToTable(typedWorkbook, sourceTable, targetTable, args)
        }

        case 'apply_plan': {
          const planId = typeof args.plan_id === 'string' ? args.plan_id.trim() : ''
          if (!planId) {
            return { error: 'plan_id is required.' }
          }

          const plan = this.planStore.get(planId)
          if (!plan) {
            return { error: `Plan not found: ${planId}` }
          }

          const latestSnapshot = this.schemaCache.refresh(typedWorkbook)

          if (plan.type === 'set_table_cell_value') {
            const table = latestSnapshot.tables.find((entry) => entry.sheetName === plan.sheetName && entry.range === plan.tableRange)
            if (!table) {
              return { error: 'The target table no longer exists. Please preview again.' }
            }

            const location = {
              sheet_name: plan.sheetName,
              target_cell: plan.targetCell,
              current_value: plan.currentValue as CellValue,
              current_formula: plan.currentFormula,
            }
            const writeResult = this.executor.setTableCellValue(typedWorkbook, location, plan.nextValue)
            if (!('error' in writeResult)) {
              this.planStore.delete(plan.id)
            }
            return writeResult
          }

          const sourceTable = latestSnapshot.tables.find((entry) =>
            entry.sheetName === plan.sourceSheetName && entry.range === plan.sourceTableRange,
          )
          const targetTable = latestSnapshot.tables.find((entry) =>
            entry.sheetName === plan.targetSheetName && entry.range === plan.targetTableRange,
          )
          if (!sourceTable || !targetTable) {
            return { error: 'The source or target table no longer exists. Please preview again.' }
          }
          if (this.getTableSignature(sourceTable) !== plan.sourceSignature || this.getTableSignature(targetTable) !== plan.targetSignature) {
            return { error: 'The table structure changed after preview. Please preview again before applying.' }
          }

          const writeResult = this.executor.syncTableToTable(typedWorkbook, sourceTable, targetTable, plan.args)
          if (!('error' in writeResult)) {
            this.planStore.delete(plan.id)
          }
          return writeResult
        }

        case 'append_table_records': {
          const targetTable = resolveTable(snapshot, args)
          if (!targetTable) {
            return { error: 'No target semantic table was detected. Prepare a target sheet with headers first.' }
          }

          return this.executor.appendTableRecords(typedWorkbook, targetTable, args.records)
        }

        case 'set_cell_value': {
          const sheet = getSheet()
          const range = sheet.getRange(args.range as string)
          this.writeScalarToRange(range, args.value)
          this.schemaCache.invalidate(typedWorkbook)
          return { success: true, message: `Set ${args.range}` }
        }

        case 'set_range_values': {
          const sheet = getSheet()
          const range = sheet.getRange(args.range as string)
          const values = args.values as (string | number | boolean | null)[][]
          const targetRange = range.getRange()
          const startRow = targetRange.startRow
          const startColumn = targetRange.startColumn
          const expectedRowCount = targetRange.endRow - targetRange.startRow + 1
          const expectedColumnCount = targetRange.endColumn - targetRange.startColumn + 1

          if (!Array.isArray(values) || values.length !== expectedRowCount) {
            return {
              error: `Range ${args.range} expects ${expectedRowCount} row(s), but received ${Array.isArray(values) ? values.length : 0}`,
            }
          }

          if (values.some((row) => !Array.isArray(row) || row.length !== expectedColumnCount)) {
            return {
              error: `Range ${args.range} expects every row to have ${expectedColumnCount} column(s)`,
            }
          }

          const plainValues = values.map((row) =>
            row.map((cell) => (isFormulaString(cell) ? null : cell ?? null)),
          )
          range.setValues(plainValues as (string | number | boolean)[][])

          values.forEach((row, rowOffset) => {
            row.forEach((cell, columnOffset) => {
              if (isFormulaString(cell)) {
                sheet.getRange(startRow + rowOffset, startColumn + columnOffset).setFormula(cell)
              }
            })
          })

          this.schemaCache.invalidate(typedWorkbook)
          return { success: true, message: `Wrote range ${args.range}` }
        }

        case 'get_range_values': {
          const sheet = getSheet()
          const range = sheet.getRange(args.range as string)
          const values = range.getValues()
          const formulas = range.getFormulas()
          return {
            range: args.range,
            values,
            formulas,
            addressed_cells: this.serializeGrid(values, formulas, range.getRow(), range.getColumn()),
          }
        }

        case 'format_range': {
          const sheet = getSheet()
          const resolvedRange = this.resolveRangeA1(sheet, args)
          if (typeof resolvedRange !== 'string') {
            return resolvedRange
          }

          const range = sheet.getRange(resolvedRange)

          this.applyRangeFormatting(range, args)
          this.schemaCache.invalidate(typedWorkbook)
          return { success: true, message: `Formatted ${resolvedRange}`, range: resolvedRange }
        }

        case 'set_freeze': {
          const sheet = getSheet()
          const frozenRows = typeof args.frozen_rows === 'number' ? Math.floor(args.frozen_rows) : 0
          const frozenColumns = typeof args.frozen_columns === 'number' ? Math.floor(args.frozen_columns) : 0

          if (frozenRows < 0 || frozenRows > 2000 || frozenColumns < 0 || frozenColumns > 2000) {
            return { error: 'frozen_rows and frozen_columns must be integers between 0 and 2000.' }
          }

          if (!sheet.setFrozenRows || !sheet.setFrozenColumns) {
            return { error: 'The current Univer API does not support setting frozen rows or columns.' }
          }

          sheet.setFrozenRows(frozenRows)
          sheet.setFrozenColumns(frozenColumns)
          this.schemaCache.invalidate(typedWorkbook)
          return {
            success: true,
            message: `Set frozen panes to ${frozenRows} row(s) and ${frozenColumns} column(s)`,
            frozen_rows: frozenRows,
            frozen_columns: frozenColumns,
          }
        }

        case 'hide_rows':
        case 'show_rows': {
          const sheet = getSheet() as SheetLike & {
            hideRows?: (rowIndex: number, count?: number) => void
            showRows?: (rowIndex: number, count?: number) => void
            unhideRow?: (row: unknown) => void
          }
          const rowIndex = args.row_index as number
          const count = typeof args.count === 'number' ? Math.max(1, Math.floor(args.count)) : 1

          if (!Number.isInteger(rowIndex) || rowIndex < 0) {
            return { error: 'row_index must be a zero-based non-negative integer.' }
          }

          if (name === 'hide_rows') {
            if (!sheet.hideRows) {
              return { error: 'The current Univer API does not support hiding rows.' }
            }
            sheet.hideRows(rowIndex, count)
          } else if (sheet.showRows) {
            sheet.showRows(rowIndex, count)
          } else if (sheet.unhideRow) {
            const lastColumn = Math.max(0, sheet.getLastColumn())
            sheet.unhideRow(sheet.getRange(rowIndex, 0, count, lastColumn + 1))
          } else {
            return { error: 'The current Univer API does not support showing hidden rows.' }
          }

          this.schemaCache.invalidate(typedWorkbook)
          return {
            success: true,
            message: `${name === 'hide_rows' ? 'Hid' : 'Showed'} ${count} row(s) starting at ${rowIndex + 1}`,
            row_index: rowIndex,
            row_number: rowIndex + 1,
            count,
          }
        }

        case 'hide_columns':
        case 'show_columns': {
          const sheet = getSheet() as SheetLike & {
            hideColumns?: (columnIndex: number, count?: number) => void
            showColumns?: (columnIndex: number, count?: number) => void
            unhideColumn?: (column: unknown) => void
          }
          const columnIndex = args.column_index as number
          const count = typeof args.count === 'number' ? Math.max(1, Math.floor(args.count)) : 1

          if (!Number.isInteger(columnIndex) || columnIndex < 0) {
            return { error: 'column_index must be a zero-based non-negative integer.' }
          }

          if (name === 'hide_columns') {
            if (!sheet.hideColumns) {
              return { error: 'The current Univer API does not support hiding columns.' }
            }
            sheet.hideColumns(columnIndex, count)
          } else if (sheet.showColumns) {
            sheet.showColumns(columnIndex, count)
          } else if (sheet.unhideColumn) {
            const lastRow = Math.max(0, sheet.getLastRow())
            sheet.unhideColumn(sheet.getRange(0, columnIndex, lastRow + 1, count))
          } else {
            return { error: 'The current Univer API does not support showing hidden columns.' }
          }

          this.schemaCache.invalidate(typedWorkbook)
          return {
            success: true,
            message: `${name === 'hide_columns' ? 'Hid' : 'Showed'} ${count} column(s) starting at ${colIndexToLetter(columnIndex)}`,
            column_index: columnIndex,
            column_letter: colIndexToLetter(columnIndex),
            count,
          }
        }

        case 'set_gridlines': {
          const sheet = getSheet()
          const hidden = args.hidden === true

          if (!sheet.setHiddenGridlines) {
            return { error: 'The current Univer API does not support setting gridline visibility.' }
          }

          sheet.setHiddenGridlines(hidden)
          this.schemaCache.invalidate(typedWorkbook)
          return {
            success: true,
            message: `${hidden ? 'Hid' : 'Showed'} gridlines`,
            hidden,
          }
        }

        case 'set_sheet_tab_color': {
          const sheet = getSheet()
          const color = typeof args.color === 'string' ? args.color.trim() : ''
          if (!color) {
            return { error: 'color is required.' }
          }

          if (!sheet.setTabColor) {
            return { error: 'The current Univer API does not support setting sheet tab color.' }
          }

          sheet.setTabColor(color)
          this.schemaCache.invalidate(typedWorkbook)
          return {
            success: true,
            message: `Set sheet tab color to ${color}`,
            color,
          }
        }

        case 'set_column_width': {
          const sheet = getSheet()
          const columnIndex = args.column_index as number
          const width = args.width as number

          if (!Number.isInteger(columnIndex) || columnIndex < 0) {
            return { error: 'column_index must be a zero-based non-negative integer.' }
          }

          if (!Number.isFinite(width) || width < 20 || width > 2000) {
            return { error: 'width must be a number between 20 and 2000 pixels.' }
          }

          sheet.setColumnWidth(columnIndex, width)
          this.schemaCache.invalidate(typedWorkbook)
          return {
            success: true,
            message: `Set column ${colIndexToLetter(columnIndex)} width to ${width}px`,
            column_index: columnIndex,
            column_letter: colIndexToLetter(columnIndex),
            width,
          }
        }

        case 'set_row_height': {
          const sheet = getSheet()
          const resolvedRows = this.resolveRowSpan(sheet, args)
          if ('error' in resolvedRows) {
            return resolvedRows
          }
          const rowIndex = resolvedRows.rowIndex
          const count = resolvedRows.count
          const height = args.height as number
          const forced = args.forced !== false

          if (!Number.isInteger(rowIndex) || rowIndex < 0) {
            return { error: 'row_index must be a zero-based non-negative integer.' }
          }

          if (!Number.isFinite(height) || height < 8 || height > 500) {
            return { error: 'height must be a number between 8 and 500 pixels.' }
          }

          if (forced && sheet.setRowHeightsForced) {
            sheet.setRowHeightsForced(rowIndex, count, height)
          } else if ('setRowHeights' in sheet && typeof sheet.setRowHeights === 'function') {
            sheet.setRowHeights(rowIndex, count, height)
          } else {
            for (let offset = 0; offset < count; offset += 1) {
              sheet.setRowHeight(rowIndex + offset, height)
            }
          }

          const rowHeights = Array.from({ length: count }, (_, offset) => ({
            row_index: rowIndex + offset,
            row_number: rowIndex + offset + 1,
            height: this.readRowHeight(sheet as unknown as SheetLike, rowIndex + offset),
          }))
          const mismatchedRows = rowHeights.filter((row) => Math.abs(row.height - height) > 1)

          this.schemaCache.invalidate(typedWorkbook)
          return {
            success: true,
            message: `Set ${count} row(s) starting at ${rowIndex + 1} height to ${height}px`,
            range: resolvedRows.range,
            row_index: rowIndex,
            row_number: rowIndex + 1,
            count,
            height,
            forced,
            row_heights: rowHeights,
            warning: mismatchedRows.length > 0
              ? `Read-back height differs on ${mismatchedRows.length} row(s). Univer may still be auto-fitting taller content.`
              : null,
          }
        }

        case 'auto_resize_columns': {
          const sheet = getSheet()
          const result = this.autoResizeColumnsForRange(sheet, args)
          if ('error' in result) {
            return result
          }

          this.schemaCache.invalidate(typedWorkbook)
          return {
            success: true,
            message: `Auto resized columns for ${result.range}`,
            ...result,
          }
        }

        case 'auto_resize_rows': {
          const sheet = getSheet()
          const result = this.autoResizeRowsForRange(sheet, args)
          if ('error' in result) {
            return result
          }

          this.schemaCache.invalidate(typedWorkbook)
          return {
            success: true,
            message: `Auto resized rows for ${result.range}`,
            ...result,
          }
        }

        case 'clear_range': {
          const sheet = getSheet()
          const range = sheet.getRange(args.range as string)
          range.clear({
            contentsOnly: args.contents_only as boolean | undefined,
            formatOnly: args.format_only as boolean | undefined,
          })
          this.schemaCache.invalidate(typedWorkbook)
          return { success: true, message: `Cleared ${args.range}` }
        }

        case 'insert_sheet': {
          const sheet = workbook.insertSheet(args.name as string)
          this.schemaCache.invalidate(typedWorkbook)
          return { success: true, sheet_name: sheet.getSheetName() }
        }

        case 'rename_sheet': {
          const sheet = typeof args.old_name === 'string' && args.old_name.trim()
            ? workbook.getSheetByName(args.old_name)
            : workbook.getActiveSheet()
          if (!sheet) {
            return {
              error: typeof args.old_name === 'string' && args.old_name.trim()
                ? `Sheet not found: ${args.old_name}`
                : 'No active sheet',
            }
          }

          sheet.setName(args.new_name as string)
          this.schemaCache.invalidate(typedWorkbook)
          return { success: true, message: `Renamed sheet to ${args.new_name}` }
        }

        case 'insert_rows': {
          const sheet = getSheet()
          const rowIndex = args.row_index as number
          const count = (args.count as number) ?? 1
          sheet.insertRows(rowIndex, count)
          this.schemaCache.invalidate(typedWorkbook)
          return { success: true, message: `Inserted ${count} row(s) before ${rowIndex + 1}` }
        }

        case 'delete_rows': {
          const sheet = getSheet()
          const rowIndex = args.row_index as number
          const count = (args.count as number) ?? 1
          sheet.deleteRows(rowIndex, count)
          this.schemaCache.invalidate(typedWorkbook)
          return { success: true, message: `Deleted ${count} row(s) from ${rowIndex + 1}` }
        }

        case 'insert_columns': {
          const sheet = getSheet()
          const columnIndex = args.column_index as number
          const count = (args.count as number) ?? 1
          sheet.insertColumns(columnIndex, count)
          this.schemaCache.invalidate(typedWorkbook)
          return { success: true, message: `Inserted ${count} column(s) before ${colIndexToLetter(columnIndex)}` }
        }

        case 'delete_columns': {
          const sheet = getSheet()
          const columnIndex = args.column_index as number
          const count = (args.count as number) ?? 1
          sheet.deleteColumns(columnIndex, count)
          this.schemaCache.invalidate(typedWorkbook)
          return { success: true, message: `Deleted ${count} column(s) from ${colIndexToLetter(columnIndex)}` }
        }

        case 'merge_cells': {
          const sheet = getSheet()
          const range = sheet.getRange(args.range as string)
          if (args.unmerge) {
            range.breakApart()
            this.schemaCache.invalidate(typedWorkbook)
            return { success: true, message: `Unmerged ${args.range}` }
          }

          const values = range.getValues()
          const formulas = range.getFormulas()
          const nonEmptyCells: string[] = []
          const bounds = range.getRange() ?? this.parseA1Bounds(args.range as string)
          values.forEach((row, rowIndex) => {
            row.forEach((value, columnIndex) => {
              const formula = formulas[rowIndex]?.[columnIndex]
              const hasValue = value !== null && value !== undefined && value !== ''
              if (hasValue || formula) {
                const absoluteRow = (bounds?.startRow ?? 0) + rowIndex
                const absoluteColumn = (bounds?.startColumn ?? 0) + columnIndex
                nonEmptyCells.push(`${colIndexToLetter(absoluteColumn)}${absoluteRow + 1}`)
              }
            })
          })

          range.merge()
          this.schemaCache.invalidate(typedWorkbook)
          return {
            success: true,
            message: `Merged ${args.range}`,
            warning: nonEmptyCells.length > 1
              ? `Range ${args.range} contained ${nonEmptyCells.length} non-empty cells (${nonEmptyCells.slice(0, 8).join(', ')}${nonEmptyCells.length > 8 ? ', ...' : ''}). Spreadsheet merge behavior may only keep the top-left visible value. Use "回到此次更新前" if this is not desired.`
              : null,
            non_empty_cells_before_merge: nonEmptyCells,
          }
        }

        default:
          return { error: `Unknown tool: ${name}` }
      }
    } catch (error) {
      return { error: error instanceof Error ? error.message : String(error) }
    }
  }

  private readRowConditions(value: unknown): Record<string, unknown> {
    if (!value || typeof value !== 'object' || Array.isArray(value)) {
      return {}
    }
    return value as Record<string, unknown>
  }

  private rowArrayToRecord(table: TableDescriptor, rowValues: CellValue[]): Record<string, unknown> {
    const record: Record<string, unknown> = {}
    table.columns.forEach((column) => {
      record[column.header] = rowValues[column.relativeColumnIndex]
    })
    return record
  }
}

function safeFiniteNumber(value: unknown, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value) ? value : fallback
}
