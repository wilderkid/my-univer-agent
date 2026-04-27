import { getUniverAPI } from './univer-ref'

export interface UniverAgentConfig {
  apiKey: string
  model: string
  baseURL: string
}

// ── OpenAI-compatible message types ──────────────────────────────────────────

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

// ── Tool definitions (OpenAI function-calling format) ────────────────────────

const TOOLS = [
  {
    type: 'function',
    function: {
      name: 'get_spreadsheet_info',
      description: '获取电子表格信息：活动工作表名称、所有工作表名称、数据范围以及当前可见单元格内容。在执行操作前先调用此函数了解表格状态。',
      parameters: { type: 'object', properties: {}, required: [] },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_cell_value',
      description: '在单元格或区域写入值或公式。值以"="开头时自动识别为公式。',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string', description: 'A1 表示法，例如 "A1" 或 "B2:D5"' },
          value: { description: '要写入的值（数字/字符串/布尔）或公式（以 = 开头的字符串）' },
          sheet_name: { type: 'string', description: '工作表名称（省略则使用活动工作表）' },
        },
        required: ['range', 'value'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_range_values',
      description: '用二维数组批量写入区域值，支持混合普通值与公式（公式以"="开头）。数组维度必须与区域一致。',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string', description: 'A1 表示法区域，例如 "A1:D4"' },
          values: {
            type: 'array',
            description: '二维数组，行 × 列，公式以"="开头。示例：[["姓名","分数"],["张三",90]]',
            items: { type: 'array', items: {} },
          },
          sheet_name: { type: 'string' },
        },
        required: ['range', 'values'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'get_range_values',
      description: '读取区域单元格的值和公式',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string', description: 'A1 表示法' },
          sheet_name: { type: 'string' },
        },
        required: ['range'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'format_range',
      description: '对区域设置格式：加粗、斜体、字号、字体颜色、背景色、下划线、水平对齐',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string' },
          bold: { type: 'boolean' },
          italic: { type: 'boolean' },
          font_size: { type: 'number' },
          font_color: { type: 'string', description: 'CSS 颜色，例如 "#ff0000" 或 "red"' },
          background: { type: 'string', description: 'CSS 背景色' },
          underline: { type: 'boolean' },
          horizontal_alignment: { type: 'string', enum: ['left', 'center', 'right'] },
          sheet_name: { type: 'string' },
        },
        required: ['range'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'clear_range',
      description: '清空区域内容和/或格式',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string' },
          contents_only: { type: 'boolean', description: '仅清除内容，保留格式' },
          format_only: { type: 'boolean', description: '仅清除格式，保留内容' },
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
      description: '新建一个工作表',
      parameters: {
        type: 'object',
        properties: {
          name: { type: 'string', description: '新工作表名称' },
        },
        required: ['name'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'rename_sheet',
      description: '重命名工作表',
      parameters: {
        type: 'object',
        properties: {
          new_name: { type: 'string' },
          old_name: { type: 'string', description: '省略则重命名当前活动工作表' },
        },
        required: ['new_name'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'insert_rows',
      description: '在指定行位置之前插入若干行（0-based 行索引）',
      parameters: {
        type: 'object',
        properties: {
          row_index: { type: 'number', description: '0-based 行索引' },
          count: { type: 'number', description: '插入行数，默认 1' },
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
      description: '从指定行位置开始删除若干行（0-based 行索引）',
      parameters: {
        type: 'object',
        properties: {
          row_index: { type: 'number' },
          count: { type: 'number', description: '删除行数，默认 1' },
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
      description: '在指定列位置之前插入若干列（0-based 列索引）',
      parameters: {
        type: 'object',
        properties: {
          column_index: { type: 'number' },
          count: { type: 'number', description: '插入列数，默认 1' },
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
      description: '从指定列位置开始删除若干列（0-based 列索引）',
      parameters: {
        type: 'object',
        properties: {
          column_index: { type: 'number' },
          count: { type: 'number', description: '删除列数，默认 1' },
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
      description: '合并或拆分单元格区域',
      parameters: {
        type: 'object',
        properties: {
          range: { type: 'string' },
          unmerge: { type: 'boolean', description: 'true 表示拆分合并单元格，默认合并' },
          sheet_name: { type: 'string' },
        },
        required: ['range'],
      },
    },
  },
]

const SYSTEM_PROMPT = `你是一个控制电子表格（Univer）的 AI 助手。
必须通过调用工具来执行操作，不要只用文字描述——直接执行。
- 行/列索引从 0 开始（第1行 = row_index 0，A列 = column_index 0）
- A1 表示法中行号从 1 开始（A1、B3、C1:D5）
- 对于复杂任务，先调用 get_spreadsheet_info 了解当前状态
- 执行后给出简洁的中文结果说明
`

// ── Column index <-> letter conversion ───────────────────────────────────────

function colIndexToLetter(index: number): string {
  let letter = ''
  let n = index
  while (n >= 0) {
    letter = String.fromCharCode(65 + (n % 26)) + letter
    n = Math.floor(n / 26) - 1
  }
  return letter
}

// ── UniverAgent ───────────────────────────────────────────────────────────────

export class UniverAgent {
  private config: UniverAgentConfig

  constructor(config: UniverAgentConfig) {
    this.config = config
  }

  async execute(instruction: string): Promise<string> {
    const api = getUniverAPI()
    if (!api) throw new Error('Univer API 未初始化')

    const context = this._getContext(api)

    const messages: Message[] = [
      { role: 'system', content: SYSTEM_PROMPT + '\n\n当前表格状态：\n' + context },
      { role: 'user', content: instruction },
    ]

    const MAX_TURNS = 10
    for (let turn = 0; turn < MAX_TURNS; turn++) {
      const { content, tool_calls } = await this._callLLM(messages)

      if (!tool_calls || tool_calls.length === 0) {
        return content ?? '操作完成'
      }

      messages.push({ role: 'assistant', content, tool_calls })

      for (const tc of tool_calls) {
        let args: Record<string, unknown> = {}
        try { args = JSON.parse(tc.function.arguments) } catch {
          messages.push({ role: 'tool', tool_call_id: tc.id, content: JSON.stringify({ error: `工具参数解析失败: ${tc.function.arguments}` }) })
          continue
        }

        const result = this._executeTool(tc.function.name, args)
        messages.push({ role: 'tool', tool_call_id: tc.id, content: JSON.stringify(result) })
      }
    }
    return '操作完成（已达到最大轮次）'
  }

  // ── Context builder ─────────────────────────────────────────────────────────

  private _getContext(api: ReturnType<typeof getUniverAPI>): string {
    const wb = api!.getActiveWorkbook()
    if (!wb) return '无工作簿'

    const sheets = wb.getSheets()
    const active = wb.getActiveSheet()
    const sheetNames = sheets.map(s => s.getSheetName()).join(', ')
    let ctx = `工作表列表：${sheetNames}\n当前活动工作表：${active?.getSheetName() ?? '未知'}\n`

    if (active) {
      const lastRow = active.getLastRow()
      const lastCol = active.getLastColumn()
      const rows = Math.min(lastRow + 1, 30)
      const cols = Math.min(lastCol + 1, 15)

      if (rows > 0 && cols > 0) {
        const notation = `A1:${colIndexToLetter(cols - 1)}${rows}`
        ctx += `\n数据区域（${notation}）：\n`
        const vals = active.getRange(notation).getValues()
        vals.forEach((row, ri) => {
          const cells = row
            .map((v, ci) => v !== null && v !== '' ? `${colIndexToLetter(ci)}${ri + 1}:${JSON.stringify(v)}` : '')
            .filter(Boolean)
          if (cells.length) ctx += cells.join('  ') + '\n'
        })
      } else {
        ctx += '（表格为空）\n'
      }
    }
    return ctx
  }

  // ── LLM call ────────────────────────────────────────────────────────────────

  private async _callLLM(messages: Message[]): Promise<{ content: string | null; tool_calls?: ToolCall[] }> {
    const url = this.config.baseURL.replace(/\/$/, '') + '/chat/completions'
    const res = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${this.config.apiKey}`,
      },
      body: JSON.stringify({
        model: this.config.model,
        messages,
        tools: TOOLS,
        tool_choice: 'auto',
        stream: false,
      }),
    })

    if (!res.ok) {
      const text = await res.text()
      throw new Error(`LLM API ${res.status}: ${text.slice(0, 200)}`)
    }

    const data = await res.json()
    const msg = data.choices?.[0]?.message
    if (!msg) throw new Error('LLM 返回格式异常')
    return { content: msg.content ?? null, tool_calls: msg.tool_calls }
  }

  // ── Tool executor ────────────────────────────────────────────────────────────

  private _executeTool(name: string, args: Record<string, unknown>): unknown {
    const api = getUniverAPI()
    if (!api) return { error: 'Univer API 未就绪' }

    const wb = api.getActiveWorkbook()
    if (!wb) return { error: '无活动工作簿' }

    const getSheet = () => {
      if (args.sheet_name) {
        const s = wb.getSheetByName(args.sheet_name as string)
        if (!s) throw new Error(`工作表不存在：${args.sheet_name}`)
        return s
      }
      const s = wb.getActiveSheet()
      if (!s) throw new Error('无活动工作表')
      return s
    }

    try {
      switch (name) {
        case 'get_spreadsheet_info': {
          const active = wb.getActiveSheet()
          const sheets = wb.getSheets()
          const lastRow = active?.getLastRow() ?? -1
          const lastCol = active?.getLastColumn() ?? -1
          const rows = Math.min(lastRow + 1, 30)
          const cols = Math.min(lastCol + 1, 15)
          let preview: unknown[][] = []
          if (active && rows > 0 && cols > 0) {
            const notation = `A1:${colIndexToLetter(cols - 1)}${rows}`
            preview = active.getRange(notation).getValues()
          }
          return {
            active_sheet: active?.getSheetName(),
            sheets: sheets.map(s => s.getSheetName()),
            last_row: lastRow,
            last_column: lastCol,
            data_preview: preview,
          }
        }

        case 'set_cell_value': {
          const sheet = getSheet()
          const range = sheet.getRange(args.range as string)
          const value = args.value
          if (typeof value === 'string' && value.startsWith('=')) {
            range.setFormula(value)
          } else {
            range.setValue(value as string | number | boolean)
          }
          return { success: true, message: `已设置 ${args.range} = ${args.value}` }
        }

        case 'set_range_values': {
          const sheet = getSheet()
          const range = sheet.getRange(args.range as string)
          const values = args.values as (string | number | boolean | null)[][]
          // Resolve the target range's top-left corner so we write at the correct position
          const targetRange = range.getRange()
          const startRow = targetRange.startRow
          const startCol = targetRange.startColumn
          // First write all non-formula values in batch via the A1-notation range
          const plainValues = values.map(row =>
            row.map(cell => (typeof cell === 'string' && cell.startsWith('=')) ? null : cell ?? null)
          )
          range.setValues(plainValues as (string | number | boolean)[][])
          // Then apply formulas individually at the correct offset (setValues treats "=xxx" as text)
          values.forEach((row, ri) => {
            row.forEach((cell, ci) => {
              if (typeof cell === 'string' && cell.startsWith('=')) {
                sheet.getRange(startRow + ri, startCol + ci).setFormula(cell)
              }
            })
          })
          return { success: true, message: `已批量写入 ${args.range}` }
        }

        case 'get_range_values': {
          const sheet = getSheet()
          const range = sheet.getRange(args.range as string)
          return {
            values: range.getValues(),
            formulas: range.getFormulas(),
          }
        }

        case 'format_range': {
          const sheet = getSheet()
          const range = sheet.getRange(args.range as string)
          if (args.bold !== undefined) range.setFontWeight(args.bold ? 'bold' : 'normal')
          if (args.italic !== undefined) range.setFontStyle(args.italic ? 'italic' : 'normal')
          if (args.font_size !== undefined) range.setFontSize(args.font_size as number)
          if (args.font_color !== undefined) range.setFontColor(args.font_color as string)
          if (args.background !== undefined) range.setBackground(args.background as string)
          if (args.underline !== undefined) range.setFontLine(args.underline ? 'underline' : 'none')
          if (args.horizontal_alignment !== undefined) {
            const alignMap: Record<string, string> = { left: 'left', center: 'center', right: 'normal' }
            range.setHorizontalAlignment((alignMap[args.horizontal_alignment as string] ?? 'left') as 'left' | 'center' | 'normal')
          }
          return { success: true, message: `已格式化 ${args.range}` }
        }

        case 'clear_range': {
          const sheet = getSheet()
          const range = sheet.getRange(args.range as string)
          range.clear({
            contentsOnly: args.contents_only as boolean | undefined,
            formatOnly: args.format_only as boolean | undefined,
          })
          return { success: true, message: `已清空 ${args.range}` }
        }

        case 'insert_sheet': {
          const sheet = wb.insertSheet(args.name as string)
          return { success: true, sheet_name: sheet.getSheetName() }
        }

        case 'rename_sheet': {
          const sheet = args.old_name
            ? wb.getSheetByName(args.old_name as string) ?? wb.getActiveSheet()
            : wb.getActiveSheet()
          if (!sheet) return { error: '工作表未找到' }
          sheet.setName(args.new_name as string)
          return { success: true, message: `已重命名为 "${args.new_name}"` }
        }

        case 'insert_rows': {
          const sheet = getSheet()
          sheet.insertRows(args.row_index as number, (args.count as number) ?? 1)
          return { success: true, message: `已在第 ${(args.row_index as number) + 1} 行前插入 ${args.count ?? 1} 行` }
        }

        case 'delete_rows': {
          const sheet = getSheet()
          sheet.deleteRows(args.row_index as number, (args.count as number) ?? 1)
          return { success: true, message: `已删除从第 ${(args.row_index as number) + 1} 行开始的 ${args.count ?? 1} 行` }
        }

        case 'insert_columns': {
          const sheet = getSheet()
          sheet.insertColumns(args.column_index as number, (args.count as number) ?? 1)
          return { success: true, message: `已在第 ${colIndexToLetter(args.column_index as number)} 列前插入 ${args.count ?? 1} 列` }
        }

        case 'delete_columns': {
          const sheet = getSheet()
          sheet.deleteColumns(args.column_index as number, (args.count as number) ?? 1)
          return { success: true, message: `已删除从 ${colIndexToLetter(args.column_index as number)} 列开始的 ${args.count ?? 1} 列` }
        }

        case 'merge_cells': {
          const sheet = getSheet()
          const range = sheet.getRange(args.range as string)
          if (args.unmerge) {
            range.breakApart()
            return { success: true, message: `已拆分 ${args.range}` }
          } else {
            range.merge()
            return { success: true, message: `已合并 ${args.range}` }
          }
        }

        default:
          return { error: `未知工具：${name}` }
      }
    } catch (e) {
      return { error: e instanceof Error ? e.message : String(e) }
    }
  }
}
