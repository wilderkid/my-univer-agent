import type { PersistedWorkbook } from './storage'

const API_BASE = import.meta.env.VITE_APP_API_BASE ?? ''
const IMPORT_TIMEOUT_MS = 120000

interface ImportXlsxWorkbookOptions {
  diagnostics?: boolean
}

export async function importXlsxWorkbook(file: File, options: ImportXlsxWorkbookOptions = {}): Promise<PersistedWorkbook> {
  const controller = new AbortController()
  const timeoutId = window.setTimeout(() => controller.abort(), IMPORT_TIMEOUT_MS)

  try {
    const response = await fetch(`${API_BASE}/api/workbook/import-xlsx`, {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Content-Type': 'application/octet-stream',
        'X-File-Name': encodeURIComponent(file.name),
        ...(options.diagnostics ? { 'X-Import-Diagnostics': '1' } : {}),
      },
      body: await file.arrayBuffer(),
      signal: controller.signal,
    })

    if (!response.ok) {
      throw new Error(await readErrorMessage(response, 'Excel 导入失败'))
    }

    return response.json() as Promise<PersistedWorkbook>
  } catch (error) {
    if (error instanceof DOMException && error.name === 'AbortError') {
      throw new Error('Excel 导入超时，请检查文件是否过大或是否包含大量空白格式区域。', { cause: error })
    }
    throw error
  } finally {
    window.clearTimeout(timeoutId)
  }
}

export async function exportXlsxWorkbook(workbook: PersistedWorkbook): Promise<Blob> {
  const response = await fetch(`${API_BASE}/api/workbook/export-xlsx`, {
    method: 'POST',
    credentials: 'include',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(workbook),
  })

  if (!response.ok) {
    throw new Error(await readErrorMessage(response, 'Excel 导出失败'))
  }

  return response.blob()
}

async function readErrorMessage(response: Response, fallback: string): Promise<string> {
  try {
    const payload = await response.json()
    return typeof payload?.error === 'string' ? payload.error : fallback
  } catch {
    return fallback
  }
}
