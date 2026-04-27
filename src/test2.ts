import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import '@univerjs/preset-sheets-core/lib/index.css'

// Create container dynamically and append to body
const container = document.createElement('div')
container.style.cssText = 'width:100vw;height:100vh;'
document.body.appendChild(container)

// Pass HTMLElement directly — avoids any getElementById timing issues
const { univerAPI } = createUniver({
  locale: LocaleType.ZH_CN,
  locales: { [LocaleType.ZH_CN]: mergeLocales(UniverPresetSheetsCoreZhCN) },
  presets: [UniverSheetsCorePreset({ container })],
})

// Delay workbook creation so Univer's React tree can finish rendering first
setTimeout(() => {
  univerAPI.createWorkbook({ name: 'Test' })
  console.log('[test2] workbook created (delayed)')

  // Pixel check 1 second later
  setTimeout(() => {
    document.querySelectorAll('canvas').forEach((c, i) => {
      const ctx = c.getContext('2d')!
      const d = ctx.getImageData(0, 0, c.width || 1, c.height || 1)
      let count = 0
      for (let j = 3; j < d.data.length; j += 4) if (d.data[j] > 0) count++
      console.log(`[test2] canvas[${i}] ${c.width}×${c.height}: ${count} non-transparent pixels`)
    })
  }, 1000)
}, 300)
