import { createRoot } from 'react-dom/client'
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import '@univerjs/preset-sheets-core/lib/index.css'
import './index.css'
import { AgentPanel } from './components/AgentPanel'
import { setUniverAPI } from './lib/univer-ref'
import { saveWorkbook, loadWorkbook, clearWorkbook, type PersistedSheet } from './lib/storage'

// ── Univer container — created dynamically ───────────────────────────────────
const univerContainer = document.createElement('div')
univerContainer.style.cssText =
  'position:fixed;top:0;left:0;width:100vw;height:100vh;z-index:0;'
document.body.appendChild(univerContainer)

const { univerAPI } = createUniver({
  locale: LocaleType.ZH_CN,
  locales: {
    [LocaleType.ZH_CN]: mergeLocales(UniverPresetSheetsCoreZhCN),
  },
  presets: [
    UniverSheetsCorePreset({ container: univerContainer }),
  ],
})

// Register univerAPI so UniverAgent can access it
setUniverAPI(univerAPI)

// ── Workbook init & auto-save ────────────────────────────────────────────────
let bootstrapped = false
let canSave = false
let dirty = false

function forceCanvasResize() {
  const fire = () => window.dispatchEvent(new Event('resize'))
  requestAnimationFrame(() => {
    fire()
    requestAnimationFrame(() => {
      fire()
      setTimeout(fire, 120)
      setTimeout(fire, 360)
    })
  })
}

function persistWorkbook() {
  if (!canSave) return
  const wb = univerAPI.getActiveWorkbook()
  if (!wb) return
  try {
    const sheets: PersistedSheet[] = wb.getSheets().map(sheet => {
      const lastRow = sheet.getLastRow()
      const lastColumn = sheet.getLastColumn()
      const hasData = lastRow >= 0 && lastColumn >= 0

      if (!hasData) {
        return {
          name: sheet.getSheetName(),
          values: [],
          formulas: [],
        }
      }

      const range = sheet.getRange(0, 0, lastRow + 1, lastColumn + 1)
      const rawValues = range.getValues()
      const rawFormulas = range.getFormulas()
      return {
        name: sheet.getSheetName(),
        values: rawValues.map(row => row.map(v => v ?? null)) as (string | number | boolean | null)[][],
        formulas: rawFormulas.map(row => row.map(f => f ?? '')),
      }
    })

    saveWorkbook({
      version: 'grid-v1',
      name: wb.getName(),
      activeSheetName: wb.getActiveSheet().getSheetName(),
      sheets,
    })
    dirty = false
  } catch (e) {
    console.warn('[univer-agent] 工作簿保存失败：', e)
  }
}

function restoreWorkbookData() {
  const saved = loadWorkbook()

  // Always create workbook — Univer auto-generates a default sheet
  const wb = univerAPI.createWorkbook({ name: saved?.name || '我的表格' })

  if (!saved || saved.sheets.length === 0) return false

  const sheets = saved.sheets
  // Rename the default first sheet to match saved data
  const firstSheet = wb.getActiveSheet()
  if (firstSheet && firstSheet.getSheetName() !== sheets[0].name) {
    firstSheet.setName(sheets[0].name)
  }

  sheets.forEach((sheetData, index) => {
    const targetSheet = index === 0
      ? firstSheet!
      : wb.insertSheet(sheetData.name)
    if (!targetSheet) return
    if (targetSheet.getSheetName() !== sheetData.name) {
      targetSheet.setName(sheetData.name)
    }

    const rowCount = sheetData.values.length
    const columnCount = Math.max(0, ...sheetData.values.map(row => row.length), ...sheetData.formulas.map(row => row.length))

    if (rowCount > 0 && columnCount > 0) {
      // Replace null with empty string for setValues compatibility (CellValue = string | number | boolean)
      const plainValues: (string | number | boolean)[][] = sheetData.values.map(row =>
        row.map(cell => (cell === null || cell === undefined) ? '' : cell as string | number | boolean)
      )
      targetSheet.getRange(0, 0, rowCount, columnCount).setValues(plainValues)

      sheetData.formulas.forEach((row, rowIndex) => {
        row.forEach((formula, colIndex) => {
          if (formula) {
            targetSheet.getRange(rowIndex, colIndex).setFormula(formula)
          }
        })
      })
    }
  })

  const activeSheet = wb.getSheetByName(saved.activeSheetName)
  if (activeSheet) {
    wb.setActiveSheet(activeSheet)
  }

  return true
}

function bootstrapWorkbook() {
  if (bootstrapped) return
  bootstrapped = true

  let restoredFromStorage = false

  try {
    restoredFromStorage = restoreWorkbookData()
  } catch (e) {
    console.warn('[univer-agent] 快照加载失败，已重置：', e)
    clearWorkbook()
    // If restore failed after workbook was already created, just reset the data
    try {
      const wb = univerAPI.getActiveWorkbook()
      if (wb) {
        const sheet = wb.getActiveSheet()
        if (sheet) {
          const lastRow = sheet.getLastRow()
          const lastCol = sheet.getLastColumn()
          if (lastRow >= 0 && lastCol >= 0) {
            sheet.getRange(0, 0, lastRow + 1, lastCol + 1).clear()
          }
        }
      }
    } catch { /* give up */ }
  }

  forceCanvasResize()
  if (restoredFromStorage) {
    setTimeout(forceCanvasResize, 500)
  }

  // 初始化完成后立刻落一次盘，确保后续刷新拿到的是当前健康快照
  setTimeout(() => {
    persistWorkbook()
    canSave = true
  }, 300)
}

univerAPI.addEvent(univerAPI.Event.LifeCycleChanged, ({ stage }) => {
  if (stage === univerAPI.Enum.LifecycleStages.Rendered || stage === univerAPI.Enum.LifecycleStages.Steady) {
    bootstrapWorkbook()
  }
})

// 如果生命周期事件因为时序原因没有接到，保留一个兜底初始化
setTimeout(() => {
  bootstrapWorkbook()
}, 1200)

// 不在命令执行回调里直接读取整张表，避免与单元格编辑提交流程互相影响
univerAPI.addEvent(univerAPI.Event.CommandExecuted, () => {
  if (!canSave) return
  dirty = true
})

window.addEventListener('pagehide', () => {
  persistWorkbook()
})

document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'hidden') {
    persistWorkbook()
  }
})

setInterval(() => {
  if (dirty) {
    persistWorkbook()
  }
}, 2500)

// ── Mount the AgentPanel overlay ─────────────────────────────────────────────
createRoot(document.getElementById('root')!).render(<AgentPanel />)
