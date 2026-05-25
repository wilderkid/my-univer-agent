import { useEffect, useRef, useState, type CSSProperties, type ReactNode } from 'react'
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import { ClearSelectionContentCommand, SetSelectionsOperation } from '@univerjs/sheets'
import type { ISelectionWithStyle } from '@univerjs/sheets'
import { IEditorBridgeService as SheetsEditorBridgeServiceIdentifier, SelectAllCommand } from '@univerjs/sheets-ui'
import type { IEditorBridgeService as SheetsEditorBridgeService } from '@univerjs/sheets-ui'
import { BreakLineCommand, IEditorService as DocsEditorServiceIdentifier } from '@univerjs/docs-ui'
import type { IEditorService as DocsEditorService } from '@univerjs/docs-ui'
import { CopyCommand, CutCommand, PasteCommand } from '@univerjs/ui'
import '@univerjs/preset-sheets-core/lib/index.css'
import { setUniverAPI } from '../lib/univer-ref'
import { clearWorkbook, getWorkbookSnapshotStats, loadWorkbook, saveWorkbook, type PersistedFileGroup, type PersistedImportSummary, type PersistedIndexRange, type PersistedSheet, type PersistedSheetView, type PersistedWorkbook, type WorkbookStorageResult } from '../lib/storage'
import { exportXlsxWorkbook, importXlsxWorkbook } from '../lib/workbook-io'
import { ChangePasswordButton } from './ChangePasswordDialog'
import { UserManualButton } from './UserManualDialog'
import { UserManagementButton } from './UserManagementDialog'
import type { AuthUser } from '../lib/auth'

interface UniverWorkspaceProps {
  overlay?: ReactNode
  currentUser: AuthUser
  onLogout: () => void
}

export function UniverWorkspace({ overlay, currentUser, onLogout }: UniverWorkspaceProps) {
  const frameRef = useRef<HTMLDivElement | null>(null)
  const hostRef = useRef<HTMLDivElement | null>(null)
  const fileInputRef = useRef<HTMLInputElement | null>(null)
  const workspaceActiveRef = useRef(false)
  const importWorkbookRef = useRef<(workbook: PersistedWorkbook) => Promise<void>>(async () => {})
  const exportWorkbookRef = useRef<() => PersistedWorkbook | null>(() => null)
  const activateFileGroupRef = useRef<(fileId: string) => void>(() => {})
  const persistedWorkbookRef = useRef<PersistedWorkbook | null>(null)
  const activeFileIdRef = useRef('')
  const fileGroupsRef = useRef<PersistedFileGroup[]>([])
  const sheetMetaRef = useRef(new Map<string, SheetSourceMeta>())
  const currentFileNameRef = useRef('我的表格.xlsx')
  const [currentFileName, setCurrentFileName] = useState('我的表格.xlsx')
  const [activeFileId, setActiveFileId] = useState('')
  const [fileGroups, setFileGroups] = useState<PersistedFileGroup[]>([])
  const [fileGroupSheetCounts, setFileGroupSheetCounts] = useState<Map<string, number>>(new Map())
  const [ioStatus, setIoStatus] = useState('支持 .xlsx 导入/导出基础数据、公式、样式、合并单元格')
  const [ioBusy, setIoBusy] = useState(false)
  const [importDiagnosticsEnabled, setImportDiagnosticsEnabled] = useState(false)

  useEffect(() => {
    currentFileNameRef.current = currentFileName
  }, [currentFileName])

  useEffect(() => {
    activeFileIdRef.current = activeFileId
  }, [activeFileId])

  useEffect(() => {
    fileGroupsRef.current = fileGroups
  }, [fileGroups])

  function setActiveFileGroupState(fileId: string, groups = fileGroupsRef.current) {
    const activeGroup = groups.find((group) => group.id === fileId)
    const nextFileId = activeGroup ? fileId : ''

    activeFileIdRef.current = nextFileId
    setActiveFileId(nextFileId)

    if (activeGroup) {
      currentFileNameRef.current = activeGroup.fileName
      setCurrentFileName(activeGroup.fileName)
    }
  }

  function syncWorkbookGroupingState(workbook: PersistedWorkbook) {
    const nextGroups = workbook.fileGroups?.length
      ? workbook.fileGroups
      : createGroupsFromSheets(workbook.sheets, workbook.fileName || currentFileNameRef.current)
    const nextActiveFileId = workbook.activeFileId && nextGroups.some((group) => group.id === workbook.activeFileId)
      ? workbook.activeFileId
      : (nextGroups[0]?.id ?? '')

    fileGroupsRef.current = nextGroups
    setFileGroups(nextGroups)
    setFileGroupSheetCounts(getFileGroupSheetCounts(workbook))
    setActiveFileGroupState(nextActiveFileId, nextGroups)

    if (!nextActiveFileId) {
      const nextFileName = normalizeFileNameInput(workbook.fileName || 'workbook.xlsx')
      currentFileNameRef.current = nextFileName
      setCurrentFileName(nextFileName)
    }
  }

  useEffect(() => {
    const frame = frameRef.current
    const host = hostRef.current
    if (!frame || !host) return

    const { univer, univerAPI } = createUniver({
      locale: LocaleType.ZH_CN,
      locales: {
        [LocaleType.ZH_CN]: mergeLocales(UniverPresetSheetsCoreZhCN),
      },
      presets: [
        UniverSheetsCorePreset({
          container: host,
          contextMenu: true,
        }),
      ],
    })

    setUniverAPI(univerAPI)

    const injector = univer.__getInjector()

    let bootstrapped = false
    let canSave = false
    let dirty = false
    let persistTimer = 0
    let fallbackTimer = 0
    let resizeTimer = 0
    let resizeTimer2 = 0
    let activeFileSyncTimer = 0

    function forceCanvasResize() {
      const fire = () => window.dispatchEvent(new Event('resize'))
      requestAnimationFrame(() => {
        fire()
        requestAnimationFrame(() => {
          fire()
          resizeTimer = window.setTimeout(fire, 120)
          resizeTimer2 = window.setTimeout(fire, 360)
        })
      })
    }

    function syncActiveFileFromWorkbook(workbook: LocalWorkbook | null | undefined) {
      if (!workbook) return
      const activeSheet = workbook.getActiveSheet()
      const meta = getSheetSourceMeta(activeSheet, sheetMetaRef.current, persistedWorkbookRef.current)
      const nextFileId = meta?.sourceFileId ?? ''

      if (!nextFileId || nextFileId === activeFileIdRef.current) return
      if (!fileGroupsRef.current.some((group) => group.id === nextFileId)) return

      setActiveFileGroupState(nextFileId)
    }

    function scheduleActiveFileSync(delay = 80) {
      window.clearTimeout(activeFileSyncTimer)
      activeFileSyncTimer = window.setTimeout(() => {
        syncActiveFileFromWorkbook(univerAPI.getActiveWorkbook() as unknown as LocalWorkbook | null)
      }, delay)
    }

    async function persistWorkbook() {
      if (!canSave) return
      const workbook = univerAPI.getActiveWorkbook()
      if (!workbook) return

      try {
        const capturedSnapshot = normalizePersistedWorkbook(capturePersistedWorkbook(
          workbook as unknown as LocalWorkbook,
          currentFileNameRef.current,
          fileGroupsRef.current,
          activeFileIdRef.current,
          sheetMetaRef.current,
        ))
        const persistedSnapshot = persistedWorkbookRef.current
          ? normalizePersistedWorkbook(persistedWorkbookRef.current)
          : null
        const removedPersistedSheets = persistedSnapshot
          ? getRemovedPersistedSheets(persistedSnapshot, capturedSnapshot)
          : []
        const snapshot = persistedSnapshot
          ? mergePersistedPresentationIntoLiveWorkbook(capturedSnapshot, persistedSnapshot)
          : capturedSnapshot
        const result = await saveWorkbook(snapshot, { allowStyleLoss: removedPersistedSheets.length > 0 })
        logWorkbookStorageResult('auto-save', result)
        if (result.ok) {
          persistedWorkbookRef.current = snapshot
          syncSheetMetadataFromWorkbook(workbook as unknown as LocalWorkbook, snapshot, sheetMetaRef.current)
          syncWorkbookGroupingState(snapshot)
          dirty = false
        }
      } catch (error) {
        console.warn('[univer-agent] Failed to persist workbook:', error)
      }
    }

    async function restoreWorkbookData() {
      const saved = await loadWorkbook()

      if (!saved || saved.sheets.length === 0) {
        univerAPI.createWorkbook({ name: saved?.name || '我的表格' })
        persistedWorkbookRef.current = null
        return false
      }
      const normalizedSaved = normalizePersistedWorkbook(saved)
      persistedWorkbookRef.current = normalizedSaved
      const workbook = univerAPI.createWorkbook(createUniverSnapshotFromPersistedWorkbook(normalizedSaved))
      applyPersistedWorkbookToWorkbook(
        workbook as unknown as LocalWorkbook,
        normalizedSaved,
        univerAPI as unknown as { Enum?: Record<string, Record<string, unknown>> },
      )

      const restoredActiveSheet = workbook.getSheetByName(normalizedSaved.activeSheetName)
      if (restoredActiveSheet) {
        workbook.setActiveSheet(restoredActiveSheet)
      }

      syncWorkbookGroupingState(normalizedSaved)
      syncSheetMetadataFromWorkbook(workbook as unknown as LocalWorkbook, normalizedSaved, sheetMetaRef.current)
      setCurrentFileName(normalizedSaved.fileName || `${normalizedSaved.name || '我的表格'}.xlsx`)
      validateWorkbookRestore(workbook as unknown as LocalWorkbook, normalizedSaved)
      logWorkbookSnapshotStats('restore', normalizedSaved)
      return true

    }

    async function importPersistedWorkbook(nextWorkbook: PersistedWorkbook) {
      const normalizedNextWorkbook = normalizePersistedWorkbook(nextWorkbook)
      const workbook = univerAPI.getActiveWorkbook() ?? univerAPI.createWorkbook({ name: nextWorkbook.name || '我的表格' })
      applyPersistedWorkbookToWorkbook(workbook as unknown as LocalWorkbook, normalizedNextWorkbook, univerAPI as unknown as { Enum?: Record<string, Record<string, unknown>> })

      const activeSheet = workbook.getSheetByName(normalizedNextWorkbook.activeSheetName)
      if (activeSheet) {
        workbook.setActiveSheet(activeSheet)
      }

      const nextFileName = nextWorkbook.fileName || `${nextWorkbook.name || '我的表格'}.xlsx`
      currentFileNameRef.current = nextFileName
      setCurrentFileName(nextFileName)
      syncWorkbookGroupingState(normalizedNextWorkbook)
      syncSheetMetadataFromWorkbook(workbook as unknown as LocalWorkbook, normalizedNextWorkbook, sheetMetaRef.current)
      persistedWorkbookRef.current = normalizedNextWorkbook
      const result = await saveWorkbook(normalizedNextWorkbook, { allowStyleLoss: true })
      logWorkbookStorageResult('import-save', result)
      if (result.ok) dirty = false
      forceCanvasResize()
    }

    async function bootstrapWorkbook() {
      if (bootstrapped) return
      bootstrapped = true

      let restoredFromStorage = false
      try {
        restoredFromStorage = await restoreWorkbookData()
      } catch (error) {
        console.warn('[univer-agent] Snapshot restore failed, clearing storage:', error)
        void clearWorkbook()
      }

      forceCanvasResize()
      if (restoredFromStorage) {
        window.setTimeout(forceCanvasResize, 500)
      }

      window.setTimeout(() => {
        void persistWorkbook()
        canSave = true
      }, 300)
    }

    importWorkbookRef.current = async (nextWorkbook: PersistedWorkbook) => {
      const previousCanSave = canSave
      canSave = false
      await importPersistedWorkbook(nextWorkbook).finally(() => {
        canSave = previousCanSave
        window.setTimeout(() => {
          forceCanvasResize()
          canSave = true
        }, 300)
      })
    }

    exportWorkbookRef.current = () => {
      const workbook = univerAPI.getActiveWorkbook()
      return workbook
        ? capturePersistedWorkbook(
          workbook as unknown as LocalWorkbook,
          currentFileNameRef.current,
          fileGroupsRef.current,
          activeFileIdRef.current,
          sheetMetaRef.current,
        )
        : null
    }

    activateFileGroupRef.current = (fileId: string) => {
      const workbook = univerAPI.getActiveWorkbook()
      const snapshot = persistedWorkbookRef.current
      const targetSheetName = snapshot?.sheets.find((sheet) => sheet.sourceFileId === fileId)?.name
      const targetSheet = targetSheetName ? workbook?.getSheetByName(targetSheetName) : null

      if (targetSheet) {
        workbook?.setActiveSheet(targetSheet)
      }

      setActiveFileGroupState(fileId)
      forceCanvasResize()
    }

    univerAPI.addEvent(univerAPI.Event.LifeCycleChanged, ({ stage }) => {
      if (
        stage === univerAPI.Enum.LifecycleStages.Rendered ||
        stage === univerAPI.Enum.LifecycleStages.Steady
      ) {
        void bootstrapWorkbook()
      }
    })

    fallbackTimer = window.setTimeout(() => {
      void bootstrapWorkbook()
    }, 1200)

    univerAPI.addEvent(univerAPI.Event.CommandExecuted, () => {
      syncActiveFileFromWorkbook(univerAPI.getActiveWorkbook() as unknown as LocalWorkbook | null)
      if (!canSave) return
      dirty = true
    })

    const handlePageHide = () => {
      void persistWorkbook()
    }

    const handleVisibilityChange = () => {
      if (document.visibilityState === 'hidden') {
        void persistWorkbook()
      }
    }

    const activateWorkspace = () => {
      workspaceActiveRef.current = true
    }

    const deactivateWorkspace = () => {
      workspaceActiveRef.current = false
    }

    const handleWorkspacePointerDown = () => {
      activateWorkspace()
      scheduleActiveFileSync()
    }

    const handleWindowBlur = () => {
      deactivateWorkspace()
    }

    const handleShortcutKeyDown = (event: KeyboardEvent) => {
      const editorBridgeService = getUniverService<SheetsEditorBridgeService>(
        injector,
        SheetsEditorBridgeServiceIdentifier,
      )
      const editorService = getUniverService<DocsEditorService>(injector, DocsEditorServiceIdentifier)
      const editableTarget = isEditableElement(event.target)
      const univerEditorTarget = isUniverEditorElement(event.target)
      const cellEditorVisible = editorBridgeService?.isVisible().visible ?? false
      const focusedEditor = editorService?.getFocusEditor() ?? null
      const editingCell = cellEditorVisible && !!focusedEditor

      if (event.altKey && !accelKey(event) && event.key === 'Enter') {
        if (!editingCell || !univerEditorTarget) {
          return
        }

        event.preventDefault()
        univerAPI.syncExecuteCommand(BreakLineCommand.id, {})
        return
      }

      if (!accelKey(event) && !event.altKey && event.key === 'Delete') {
        if (editingCell && univerEditorTarget) {
          return
        }

        const workbook = univerAPI.getActiveWorkbook()
        const worksheet = workbook?.getActiveSheet()
        const selection = worksheet?.getSelection()
        const activeRanges = selection?.getActiveRangeList().map((range) => range.getRange()) ?? []

        if (!workbook || !worksheet || activeRanges.length === 0) {
          return
        }

        event.preventDefault()
        event.stopPropagation()

        const currentCell = selection?.getCurrentCell() ?? null
        let primaryAssigned = false

        const syncedSelections: ISelectionWithStyle[] = activeRanges.map((range, index) => {
          let primary = null

          if (currentCell && rangeContainsCell(range, currentCell.actualRow, currentCell.actualColumn)) {
            primary = currentCell
            primaryAssigned = true
          } else if (!primaryAssigned && index === 0) {
            primary = createPrimaryCell(range)
            primaryAssigned = true
          }

          return {
            range,
            primary,
            style: null,
          }
        })

        univerAPI.syncExecuteCommand(SetSelectionsOperation.id, {
          unitId: workbook.getId(),
          subUnitId: worksheet.getSheetId(),
          selections: syncedSelections,
        })

        univerAPI.syncExecuteCommand(ClearSelectionContentCommand.id, {
          unitId: workbook.getId(),
          subUnitId: worksheet.getSheetId(),
          ranges: activeRanges,
        })
        return
      }

      if (!workspaceActiveRef.current || event.defaultPrevented) {
        return
      }

      if (editableTarget) {
        if (univerEditorTarget) {
          return
        }

        return
      }

      const key = event.key.toLowerCase()
      const withAccelKey = accelKey(event)

      if (
        withAccelKey &&
        !event.altKey &&
        !event.shiftKey &&
        (key === 'c' || key === 'x') &&
        hasBrowserTextSelection()
      ) {
        return
      }

      if (withAccelKey && !event.altKey && !event.shiftKey && key === 's') {
        event.preventDefault()
        void persistWorkbook()
        return
      }

      if (withAccelKey && !event.altKey && !event.shiftKey && key === 'z') {
        event.preventDefault()
        void univerAPI.undo()
        return
      }

      if (
        withAccelKey &&
        !event.altKey &&
        ((event.shiftKey && key === 'z') || (!event.shiftKey && key === 'y'))
      ) {
        event.preventDefault()
        void univerAPI.redo()
        return
      }

      if (withAccelKey && !event.altKey && !event.shiftKey && key === 'a') {
        event.preventDefault()
        void univerAPI.executeCommand(SelectAllCommand.id, {})
        return
      }

      if (withAccelKey && !event.altKey && !event.shiftKey && key === 'c') {
        event.preventDefault()
        void univerAPI.executeCommand(CopyCommand.id)
        return
      }

      if (withAccelKey && !event.altKey && !event.shiftKey && key === 'x') {
        event.preventDefault()
        void univerAPI.executeCommand(CutCommand.id)
        return
      }

      if (withAccelKey && !event.altKey && !event.shiftKey && key === 'v') {
        event.preventDefault()
        void univerAPI.executeCommand(PasteCommand.id)
        return
      }
    }

    window.addEventListener('pagehide', handlePageHide)
    window.addEventListener('blur', handleWindowBlur)
    document.addEventListener('visibilitychange', handleVisibilityChange)
    document.addEventListener('keydown', handleShortcutKeyDown, true)
    frame.addEventListener('pointerdown', handleWorkspacePointerDown, true)

    persistTimer = window.setInterval(() => {
      if (dirty) {
        void persistWorkbook()
      }
    }, 2500)

    return () => {
      window.removeEventListener('pagehide', handlePageHide)
      window.removeEventListener('blur', handleWindowBlur)
      document.removeEventListener('visibilitychange', handleVisibilityChange)
      document.removeEventListener('keydown', handleShortcutKeyDown, true)
      frame.removeEventListener('pointerdown', handleWorkspacePointerDown, true)
      window.clearInterval(persistTimer)
      window.clearTimeout(fallbackTimer)
      window.clearTimeout(resizeTimer)
      window.clearTimeout(resizeTimer2)
      window.clearTimeout(activeFileSyncTimer)
      importWorkbookRef.current = async () => {}
      exportWorkbookRef.current = () => null
      activateFileGroupRef.current = () => {}
    }
  // Univer should be initialized once for the host element lifetime.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  async function handleImportFiles(files: FileList | null) {
    const selectedFiles = Array.from(files ?? []).filter((file) => /\.xlsx$/i.test(file.name))
    if (selectedFiles.length === 0) return

    try {
      setIoBusy(true)
      const liveWorkbook = exportWorkbookRef.current()
      const currentWorkbook = chooseWorkbookForAppend(liveWorkbook, persistedWorkbookRef.current)
      let combinedWorkbook = currentWorkbook
        ? normalizePersistedWorkbook(currentWorkbook)
        : createEmptyPersistedWorkbook()
      const usedSheetNames = new Set(combinedWorkbook.sheets.map((sheet) => sheet.name))
      let latestGroup: PersistedFileGroup | null = null
      const importSummary = createEmptyImportSummary()

      for (const [index, file] of selectedFiles.entries()) {
        setIoStatus(`正在导入 ${index + 1}/${selectedFiles.length}: ${file.name} ...`)
        const importedRawWorkbook = await importXlsxWorkbook(file, { diagnostics: importDiagnosticsEnabled })
        addImportSummary(importSummary, importedRawWorkbook.importSummary)
        const importedWorkbook = normalizePersistedWorkbook(importedRawWorkbook)
        const group = createFileGroup(file.name)
        latestGroup = group
        combinedWorkbook = appendImportedWorkbook(combinedWorkbook, importedWorkbook, group, usedSheetNames)
      }

      if (latestGroup) {
        combinedWorkbook.activeFileId = latestGroup.id
        combinedWorkbook.fileName = latestGroup.fileName
        combinedWorkbook.activeSheetName = combinedWorkbook.sheets.find((sheet) => sheet.sourceFileId === latestGroup.id)?.name
          ?? combinedWorkbook.activeSheetName
      }

      const normalizedWorkbook = normalizePersistedWorkbook(combinedWorkbook)
      setIoStatus(`正在写入到 Univer：${normalizedWorkbook.sheets.length} 个工作表 ...`)
      await waitForNextFrame()
      await importWorkbookRef.current(normalizedWorkbook)
      setIoStatus(`已导入 ${selectedFiles.length} 个文件，工作表 ${normalizedWorkbook.sheets.length} 个；扫描 ${formatInteger(importSummary.cells)} 个单元格、${formatInteger(importSummary.styledCells)} 个样式、${formatInteger(importSummary.formulas)} 个公式、${formatInteger(importSummary.mergedRanges)} 个合并区域、${formatInteger(importSummary.hiddenRows + importSummary.hiddenColumns)} 个隐藏行列`)
    } catch (error) {
      setIoStatus(error instanceof Error ? error.message : 'Excel 导入失败')
    } finally {
      setIoBusy(false)
      if (fileInputRef.current) fileInputRef.current.value = ''
    }
  }

  async function handleExportFile() {
    const workbook = chooseWorkbookForAppend(exportWorkbookRef.current(), persistedWorkbookRef.current)
    if (!workbook) {
      setIoStatus('当前没有可导出的工作簿')
      return
    }
    const target = resolveFileOperationTarget(workbook, activeFileIdRef.current)
    const exportWorkbook = getWorkbookForFileGroup(workbook, target?.fileId ?? '')
    if (target && target.fileId !== activeFileIdRef.current) {
      setActiveFileGroupState(target.fileId)
    }

    try {
      setIoBusy(true)
      setIoStatus(`正在导出 ${exportWorkbook.fileName || currentFileName} ...`)
      const blob = await exportXlsxWorkbook(exportWorkbook)
      downloadBlob(blob, exportWorkbook.fileName || currentFileName || 'workbook.xlsx')
      setIoStatus(`已导出 ${exportWorkbook.fileName || currentFileName}`)
    } catch (error) {
      setIoStatus(error instanceof Error ? error.message : 'Excel 导出失败')
    } finally {
      setIoBusy(false)
    }
  }

  function handleActiveFileGroupChange(nextFileId: string) {
    setActiveFileGroupState(nextFileId)
    activateFileGroupRef.current(nextFileId)
  }

  function handleCloseActiveFileGroup() {
    const workbook = chooseWorkbookForAppend(exportWorkbookRef.current(), persistedWorkbookRef.current)
    const target = workbook ? resolveFileOperationTarget(workbook, activeFileIdRef.current) : null
    if (!workbook || !target) {
      return
    }

    const normalizedWorkbook = normalizePersistedWorkbook(workbook)
    if (target.fileId !== activeFileIdRef.current) {
      setActiveFileGroupState(target.fileId)
    }
    const sheetNames = target.sheets.map((sheet) => sheet.name).join('、')
    if (!window.confirm(`确定关闭文件「${target.group.fileName}」吗？\n将移除 ${target.sheets.length} 个工作表：${sheetNames}`)) {
      return
    }

    const nextWorkbook = removeFileGroupFromWorkbook(normalizedWorkbook, target.fileId)

    void importWorkbookRef.current(nextWorkbook)
    setIoStatus(nextWorkbook.fileGroups?.length
      ? `已关闭 ${target.group.fileName}，剩余 ${nextWorkbook.fileGroups.length} 个文件`
      : `已关闭 ${target.group.fileName}，文件列表已清空`)
  }

  function handleClearFileGroups() {
    if (fileGroupsRef.current.length > 0 && !window.confirm('确定清空所有已导入的 Excel 文件和工作表吗？')) {
      return
    }

    const nextWorkbook = normalizePersistedWorkbook(createEmptyPersistedWorkbook())
    void importWorkbookRef.current(nextWorkbook)
    setIoStatus('已清空文件列表，当前为一个空白工作簿')
  }

  return (
    <div ref={frameRef} style={workspaceFrameStyle} tabIndex={0}>
      <div style={fileToolbarStyle}>
        <input
          ref={fileInputRef}
          type="file"
          multiple
          accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          style={{ display: 'none' }}
          onChange={(event) => void handleImportFiles(event.target.files)}
        />
        <select
          value={activeFileId}
          onChange={(event) => handleActiveFileGroupChange(event.target.value)}
          style={fileGroupSelectStyle}
          disabled={ioBusy || fileGroups.length === 0}
          title="选择要导出的文件组"
        >
          {fileGroups.length === 0 ? (
            <option value="">当前工作簿</option>
          ) : (
            fileGroups.map((group) => (
              <option key={group.id} value={group.id}>
                {group.name} · {fileGroupSheetCounts.get(group.id) ?? 0} 表
              </option>
            ))
          )}
        </select>
        <label style={fileToolbarToggleStyle} title="开启后，导入 Excel 时会在后端控制台输出样式诊断日志">
          <input
            type="checkbox"
            checked={importDiagnosticsEnabled}
            onChange={(event) => setImportDiagnosticsEnabled(event.target.checked)}
            disabled={ioBusy}
            style={fileToolbarCheckboxStyle}
          />
          诊断日志
        </label>
        <button
          type="button"
          style={fileToolbarButtonStyle}
          onClick={() => fileInputRef.current?.click()}
          disabled={ioBusy}
        >
          导入 Excel
        </button>
        <button
          type="button"
          style={fileToolbarButtonStyle}
          onClick={() => void handleExportFile()}
          disabled={ioBusy}
        >
          导出 Excel
        </button>
        <button
          type="button"
          style={fileToolbarButtonStyle}
          onClick={handleCloseActiveFileGroup}
          disabled={ioBusy || fileGroups.length === 0 || !activeFileId}
        >
          关闭文件
        </button>
        <button
          type="button"
          style={fileToolbarDangerButtonStyle}
          onClick={handleClearFileGroups}
          disabled={ioBusy || fileGroups.length === 0}
        >
          清空
        </button>
        <UserManualButton buttonStyle={fileToolbarButtonStyle} />
        <UserManagementButton buttonStyle={fileToolbarButtonStyle} currentUser={currentUser} />
        <div style={userBadgeStyle}>
          <span>{currentUser.username}</span>
          <span style={userRoleStyle}>{currentUser.role === 'admin' ? '管理员' : '用户'}</span>
        </div>
        <ChangePasswordButton buttonStyle={fileToolbarButtonStyle} currentUser={currentUser} />
        <button
          type="button"
          style={fileToolbarButtonStyle}
          onClick={onLogout}
        >
          退出
        </button>
        <span style={fileStatusStyle}>{ioStatus}</span>
      </div>
      <div ref={hostRef} style={workspaceHostStyle} />
      {overlay ? <div style={overlayShellStyle}>{overlay}</div> : null}
    </div>
  )
}

type CellValue = string | number | boolean | null
type UniverCellData = {
  v?: string | number | boolean | null
  f?: string
  s?: Record<string, unknown>
}

interface SheetSourceMeta {
  sourceFileId: string
  sourceFileName: string
  originalSheetName: string
}

interface LocalRange {
  getValues(): unknown[][]
  getFormulas(): string[][]
  getCellDataGrid?: () => Array<Array<Record<string, unknown> | null>>
  getNumberFormats?: () => string[][]
  getA1Notation?: () => string
  setValues(values: Array<Array<string | number | boolean | UniverCellData>>): void
  setFormula(value: string): void
  setNumberFormat?: (pattern: string) => void
  clear(options?: { contentsOnly?: boolean; formatOnly?: boolean }): void
  setFontWeight(value: string): void
  setFontStyle(value: string): void
  setFontSize(value: number): void
  setFontColor(value: string): void
  setFontFamily?(value: string): void
  setFontLine?(value: string): void
  setBackground(value: string): void
  setHorizontalAlignment(value: string): void
  setVerticalAlignment?(value: string): void
  setTextRotation?(value: number): void
  setBorder(type: unknown, style: unknown, color?: string): void
  setWrapStrategy(strategy: unknown): void
  breakApart(): void
  merge(options?: unknown): void
}

interface LocalSheet {
  getSheetId(): string
  getSheetName(): string
  getSheet?(): {
    getRowCount?(): number
    getColumnCount?(): number
    setRowCount?(count: number): void
    setColumnCount?(count: number): void
    getHiddenRows?(start?: number, end?: number): Array<Record<string, unknown>>
    getHiddenCols?(start?: number, end?: number): Array<Record<string, unknown>>
    getFreeze?(): Record<string, unknown>
    hasHiddenGridlines?(): boolean
    getGridlinesColor?(): string | undefined
    getTabColor?(): string | undefined
  }
  getLastRow(): number
  getLastColumn(): number
  getRange(...args: [string] | [number, number] | [number, number, number, number]): LocalRange
  setRowCount?(count: number): LocalSheet | void
  setColumnCount?(count: number): LocalSheet | void
  getColumnWidth(columnIndex: number): number
  setColumnWidth(columnIndex: number, width: number): void
  getRowHeight(rowIndex: number): number
  setRowHeight(rowIndex: number, height: number): void
  hideRows?(rowIndex: number, numRows?: number): void
  hideColumns?(columnIndex: number, numColumns?: number): void
  unhideRow?(row: LocalRange): void
  unhideColumn?(column: LocalRange): void
  setFrozenRows?(rows: number): void
  setFrozenColumns?(columns: number): void
  getFrozenRows?(): number
  getFrozenColumns?(): number
  hasHiddenGridLines?(): boolean
  hasHiddenGridlines?(): boolean
  setHiddenGridlines?(hidden: boolean): void
  getGridLinesColor?(): string | undefined
  setTabColor?(color: string): void
  getTabColor?(): string | undefined
  getMergedRanges(): LocalRange[]
  setName(name: string): void
}

interface LocalWorkbook {
  getName(): string
  getSheets(): LocalSheet[]
  getActiveSheet(): LocalSheet
  getSheetByName(name: string): LocalSheet | null
  insertSheet(name: string): LocalSheet
  deleteSheet(sheet: LocalSheet | string): boolean
  setActiveSheet(sheet: LocalSheet): void
}

const MIN_VISIBLE_ROW_COUNT = 40
const MIN_VISIBLE_COLUMN_COUNT = 100
const DEFAULT_ROW_HEIGHT = 24
const DEFAULT_COLUMN_WIDTH = 88
const DIMENSION_EPSILON = 0.01

function capturePersistedWorkbook(
  workbook: LocalWorkbook,
  fileName: string,
  fileGroups: PersistedFileGroup[] = [],
  activeFileId = '',
  sheetMeta = new Map<string, SheetSourceMeta>(),
): PersistedWorkbook {
  const sheets: PersistedSheet[] = workbook.getSheets().map((sheet) => capturePersistedSheet(sheet, sheetMeta))
  const normalizedFileName = normalizeFileNameInput(fileName || `${workbook.getName() || 'workbook'}.xlsx`)

  return {
    version: 'grid-v1',
    name: stripExcelExtension(normalizedFileName) || workbook.getName() || 'Workbook',
    fileName: normalizedFileName,
    activeFileId,
    fileGroups,
    activeSheetName: workbook.getActiveSheet().getSheetName(),
    sheets,
  }
}

function validateWorkbookRestore(workbook: LocalWorkbook, snapshot: PersistedWorkbook): void {
  const restoredSheetNames = new Set(workbook.getSheets().map((sheet) => sheet.getSheetName()))
  const missingSheets = snapshot.sheets
    .map((sheet) => sheet.name)
    .filter((sheetName) => !restoredSheetNames.has(sheetName))

  if (missingSheets.length > 0 || workbook.getSheets().length !== snapshot.sheets.length) {
    console.warn('[univer-agent] Snapshot restore shape mismatch:', {
      expectedSheets: snapshot.sheets.map((sheet) => sheet.name),
      restoredSheets: [...restoredSheetNames],
      missingSheets,
    })
  }
}

function createUniverSnapshotFromPersistedWorkbook(workbook: PersistedWorkbook): Record<string, unknown> {
  const sheetEntries = workbook.sheets.map((sheet, index) => {
    const sheetId = sheet.sheetId || `sheet-${index + 1}`
    return [sheetId, createUniverWorksheetSnapshot(sheet, sheetId)] as const
  })

  return {
    id: `workbook-${Date.now()}`,
    name: workbook.name || 'Workbook',
    appVersion: '0.0.0',
    locale: LocaleType.ZH_CN,
    styles: {},
    sheetOrder: sheetEntries.map(([sheetId]) => sheetId),
    sheets: Object.fromEntries(sheetEntries),
  }
}

function createUniverWorksheetSnapshot(sheet: PersistedSheet, sheetId: string): Record<string, unknown> {
  const trimmedSheet = trimPersistedSheetBounds(sheet)
  const values = normalizeMatrix<CellValue>(trimmedSheet.values)
  const formulas = normalizeMatrix<string>(trimmedSheet.formulas)
  const styles = normalizeMatrix<Record<string, unknown> | null>(trimmedSheet.styles)
  const effectiveRowCount = Math.max(
    values.length,
    formulas.length,
    styles.length,
    getMaxIndexFromRecord(trimmedSheet.rowHeights) + 1,
    getMaxCoveredIndexFromRanges(trimmedSheet.hiddenRows) + 1,
    getMaxCoveredIndexFromMergedRanges(trimmedSheet.mergedRanges, 'row') + 1,
    0,
  )
  const effectiveColumnCount = Math.max(
    0,
    ...values.map((row) => row.length),
    ...formulas.map((row) => row.length),
    ...styles.map((row) => row.length),
    getMaxIndexFromRecord(trimmedSheet.columnWidths) + 1,
    getMaxCoveredIndexFromRanges(trimmedSheet.hiddenColumns) + 1,
    getMaxCoveredIndexFromMergedRanges(trimmedSheet.mergedRanges, 'column') + 1,
  )
  const rowCount = Math.max(MIN_VISIBLE_ROW_COUNT, effectiveRowCount)
  const columnCount = Math.max(MIN_VISIBLE_COLUMN_COUNT, effectiveColumnCount)
  const sheetView = normalizeSheetView(sheet.sheetView)

  return {
    id: sheetId,
    name: trimmedSheet.name,
    tabColor: sheetView.tabColor || '',
    hidden: 0,
    freeze: {
      xSplit: sheetView.frozenColumns ?? 0,
      ySplit: sheetView.frozenRows ?? 0,
      startRow: sheetView.frozenRows ?? 0,
      startColumn: sheetView.frozenColumns ?? 0,
    },
    rowCount,
    columnCount,
    defaultColumnWidth: DEFAULT_COLUMN_WIDTH,
    defaultRowHeight: DEFAULT_ROW_HEIGHT,
    mergeData: createUniverMergeData(trimmedSheet.mergedRanges),
    cellData: createUniverCellDataObject(values, formulas, styles, rowCount, columnCount),
    rowData: createUniverRowData(trimmedSheet.rowHeights, trimmedSheet.hiddenRows),
    columnData: createUniverColumnData(trimmedSheet.columnWidths, trimmedSheet.hiddenColumns),
    rowHeader: { width: 46 },
    columnHeader: { height: 20 },
    showGridlines: sheetView.hiddenGridlines ? 0 : 1,
    gridlinesColor: sheetView.gridlinesColor,
    rightToLeft: 0,
  }
}

function createUniverCellDataObject(
  values: CellValue[][],
  formulas: string[][],
  styles: Array<Array<Record<string, unknown> | null>>,
  rowCount: number,
  columnCount: number,
): Record<number, Record<number, UniverCellData>> {
  const cellData: Record<number, Record<number, UniverCellData>> = {}

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      const value = values[rowIndex]?.[columnIndex] ?? null
      const formula = formulas[rowIndex]?.[columnIndex] ?? ''
      const style = persistedStyleToUniverCellStyle(styles[rowIndex]?.[columnIndex] ?? null)
      if (value === null && !formula && !style) continue

      cellData[rowIndex] ??= {}
      cellData[rowIndex][columnIndex] = {
        ...(formula ? { f: formula } : { v: value }),
        ...(formula && value !== null ? { v: value } : {}),
        ...(style ? { s: style } : {}),
      }
    }
  }

  return cellData
}

function createUniverRowData(
  rowHeights: Record<string, number> | undefined,
  hiddenRows: PersistedIndexRange[] | undefined,
): Record<number, Record<string, unknown>> {
  const rowData: Record<number, Record<string, unknown>> = {}

  Object.entries(rowHeights ?? {}).forEach(([rowIndex, height]) => {
    const index = Number(rowIndex)
    if (Number.isInteger(index) && Number(height) > 0) {
      rowData[index] ??= {}
      rowData[index].h = Number(height)
    }
  })

  normalizeIndexRanges(hiddenRows).forEach((range) => {
    for (let rowIndex = range.start; rowIndex <= range.end; rowIndex += 1) {
      rowData[rowIndex] ??= {}
      rowData[rowIndex].hd = 1
    }
  })

  return rowData
}

function createUniverColumnData(
  columnWidths: Record<string, number> | undefined,
  hiddenColumns: PersistedIndexRange[] | undefined,
): Record<number, Record<string, unknown>> {
  const columnData: Record<number, Record<string, unknown>> = {}

  Object.entries(columnWidths ?? {}).forEach(([columnIndex, width]) => {
    const index = Number(columnIndex)
    if (Number.isInteger(index) && Number(width) > 0) {
      columnData[index] ??= {}
      columnData[index].w = Number(width)
    }
  })

  normalizeIndexRanges(hiddenColumns).forEach((range) => {
    for (let columnIndex = range.start; columnIndex <= range.end; columnIndex += 1) {
      columnData[columnIndex] ??= {}
      columnData[columnIndex].hd = 1
    }
  })

  return columnData
}

function createUniverMergeData(ranges: string[] | undefined): Array<Record<string, number>> {
  return (ranges ?? [])
    .map((range) => parseA1Range(range))
    .filter((range): range is Record<string, number> => !!range)
}

function parseA1Range(rangeA1: string): Record<string, number> | null {
  const [startRef, endRef = startRef] = rangeA1.split(':')
  const start = parseA1Cell(startRef)
  const end = parseA1Cell(endRef)
  if (!start || !end) return null

  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startColumn: Math.min(start.column, end.column),
    endColumn: Math.max(start.column, end.column),
  }
}

function parseA1Cell(cellRef: string): { row: number, column: number } | null {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(cellRef.trim())
  if (!match) return null

  return {
    row: Number(match[2]) - 1,
    column: columnLettersToIndex(match[1]),
  }
}

function getMaxCoveredIndexFromRanges(ranges: PersistedIndexRange[] | undefined): number {
  if (!Array.isArray(ranges) || ranges.length === 0) return -1
  const maxIndex = ranges.reduce((currentMax, range) => Math.max(currentMax, Number(range?.end)), -1)
  return Number.isInteger(maxIndex) ? maxIndex : -1
}

function getMaxCoveredIndexFromMergedRanges(
  ranges: string[] | undefined,
  axis: 'row' | 'column',
): number {
  if (!Array.isArray(ranges) || ranges.length === 0) return -1
  return ranges.reduce((currentMax, rangeA1) => {
    const parsed = parseA1Range(rangeA1)
    if (!parsed) return currentMax
    return Math.max(currentMax, axis === 'row' ? parsed.endRow : parsed.endColumn)
  }, -1)
}

function columnLettersToIndex(letters: string): number {
  return letters.toUpperCase().split('').reduce((total, letter) =>
    total * 26 + letter.charCodeAt(0) - 64, 0) - 1
}

function getMaxIndexFromRecord(record: Record<string, number> | undefined): number {
  return Math.max(
    -1,
    ...Object.keys(record ?? {})
      .map((key) => Number(key))
      .filter((index) => Number.isInteger(index) && index >= 0),
  )
}

function normalizePersistedWorkbook(workbook: PersistedWorkbook): PersistedWorkbook {
  const sheets = Array.isArray(workbook.sheets) ? workbook.sheets : []
  const normalizedSheets = sheets
    .filter((sheet) => sheet && typeof sheet.name === 'string')
    .map((sheet) => trimPersistedSheetBounds({
      name: sheet.name,
      sheetId: sheet.sheetId,
      sourceFileId: sheet.sourceFileId,
      sourceFileName: sheet.sourceFileName,
      originalSheetName: sheet.originalSheetName,
      values: normalizeMatrix<CellValue>(sheet.values),
      formulas: normalizeMatrix<string>(sheet.formulas),
      styles: normalizeMatrix<Record<string, unknown> | null>(sheet.styles),
      columnWidths: isPlainRecord(sheet.columnWidths) ? sheet.columnWidths as Record<string, number> : {},
      rowHeights: isPlainRecord(sheet.rowHeights) ? sheet.rowHeights as Record<string, number> : {},
      mergedRanges: Array.isArray(sheet.mergedRanges)
        ? sheet.mergedRanges.filter((range): range is string => typeof range === 'string')
        : [],
      hiddenRows: normalizeIndexRanges(sheet.hiddenRows),
      hiddenColumns: normalizeIndexRanges(sheet.hiddenColumns),
      sheetView: normalizeSheetView(sheet.sheetView),
      importStats: isPlainRecord(sheet.importStats) ? sheet.importStats as PersistedImportSummary : undefined,
    }))

  const normalizedFileGroups = normalizeFileGroups(workbook.fileGroups, normalizedSheets)
  const normalizedActiveFileId = typeof workbook.activeFileId === 'string' &&
    normalizedFileGroups.some((group) => group.id === workbook.activeFileId)
    ? workbook.activeFileId
    : ''

  return {
    version: 'grid-v1',
    name: typeof workbook.name === 'string' && workbook.name.trim() ? workbook.name : 'Workbook',
    fileName: typeof workbook.fileName === 'string' && workbook.fileName.trim() ? workbook.fileName : 'workbook.xlsx',
    activeFileId: normalizedActiveFileId,
    fileGroups: normalizedFileGroups,
    importSummary: isPlainRecord(workbook.importSummary) ? workbook.importSummary as PersistedImportSummary : undefined,
    activeSheetName: typeof workbook.activeSheetName === 'string' && workbook.activeSheetName.trim()
      ? workbook.activeSheetName
      : (normalizedSheets[0]?.name ?? 'Sheet1'),
    sheets: normalizedSheets.length > 0
      ? normalizedSheets
      : [{ name: 'Sheet1', values: [], formulas: [], styles: [], columnWidths: {}, rowHeights: {}, mergedRanges: [] }],
  }
}

function normalizeMatrix<T>(value: T[][] | null | undefined): T[][] {
  return Array.isArray(value)
    ? value.map((row) => (Array.isArray(row) ? row : []))
    : []
}

function trimPersistedSheetBounds(sheet: PersistedSheet): PersistedSheet {
  const values = normalizeMatrix<CellValue>(sheet.values)
  const formulas = normalizeMatrix<string>(sheet.formulas)
  const styles = normalizeMatrix<Record<string, unknown> | null>(sheet.styles)
  const rowHeights = isPlainRecord(sheet.rowHeights) ? sheet.rowHeights as Record<string, number> : {}
  const columnWidths = isPlainRecord(sheet.columnWidths) ? sheet.columnWidths as Record<string, number> : {}
  const hiddenRows = normalizeIndexRanges(sheet.hiddenRows)
  const hiddenColumns = normalizeIndexRanges(sheet.hiddenColumns)
  const mergedRanges = Array.isArray(sheet.mergedRanges) ? sheet.mergedRanges.filter((range): range is string => typeof range === 'string') : []

  let maxRow = -1
  let maxColumn = -1

  for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
    const row = values[rowIndex] ?? []
    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      if (row[columnIndex] !== null && row[columnIndex] !== '') {
        maxRow = Math.max(maxRow, rowIndex)
        maxColumn = Math.max(maxColumn, columnIndex)
      }
    }
  }

  for (let rowIndex = 0; rowIndex < formulas.length; rowIndex += 1) {
    const row = formulas[rowIndex] ?? []
    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      if (typeof row[columnIndex] === 'string' && row[columnIndex].trim()) {
        maxRow = Math.max(maxRow, rowIndex)
        maxColumn = Math.max(maxColumn, columnIndex)
      }
    }
  }

  for (let rowIndex = 0; rowIndex < styles.length; rowIndex += 1) {
    const row = styles[rowIndex] ?? []
    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      if (row[columnIndex]) {
        maxRow = Math.max(maxRow, rowIndex)
        maxColumn = Math.max(maxColumn, columnIndex)
      }
    }
  }

  Object.entries(rowHeights).forEach(([rowIndex, height]) => {
    const numericIndex = Number(rowIndex)
    const numericHeight = Number(height)
    if (Number.isInteger(numericIndex) && numericIndex >= 0 && Number.isFinite(numericHeight) && Math.abs(numericHeight - DEFAULT_ROW_HEIGHT) > DIMENSION_EPSILON) {
      maxRow = Math.max(maxRow, numericIndex)
    }
  })

  Object.entries(columnWidths).forEach(([columnIndex, width]) => {
    const numericIndex = Number(columnIndex)
    const numericWidth = Number(width)
    if (Number.isInteger(numericIndex) && numericIndex >= 0 && Number.isFinite(numericWidth) && Math.abs(numericWidth - DEFAULT_COLUMN_WIDTH) > DIMENSION_EPSILON) {
      maxColumn = Math.max(maxColumn, numericIndex)
    }
  })

  hiddenRows.forEach((range) => {
    maxRow = Math.max(maxRow, range.end)
  })
  hiddenColumns.forEach((range) => {
    maxColumn = Math.max(maxColumn, range.end)
  })
  mergedRanges.forEach((rangeA1) => {
    const parsed = parseA1Range(rangeA1)
    if (!parsed) return
    maxRow = Math.max(maxRow, parsed.endRow)
    maxColumn = Math.max(maxColumn, parsed.endColumn)
  })

  const trimmedRowCount = maxRow + 1
  const trimmedColumnCount = maxColumn + 1

  const trimmedValues = trimmedRowCount > 0 && trimmedColumnCount > 0
    ? values.slice(0, trimmedRowCount).map((row) =>
      Array.from({ length: trimmedColumnCount }, (_, columnIndex) => normalizeCellValue(row?.[columnIndex] ?? null)))
    : []
  const trimmedFormulas = trimmedRowCount > 0 && trimmedColumnCount > 0
    ? formulas.slice(0, trimmedRowCount).map((row) =>
      Array.from({ length: trimmedColumnCount }, (_, columnIndex) => {
        const formula = row?.[columnIndex]
        return typeof formula === 'string' ? formula : ''
      }))
    : []
  const trimmedStyles = trimmedRowCount > 0 && trimmedColumnCount > 0
    ? styles.slice(0, trimmedRowCount).map((row) =>
      Array.from({ length: trimmedColumnCount }, (_, columnIndex) => row?.[columnIndex] ?? null))
    : []

  const trimmedRowHeights = Object.fromEntries(
    Object.entries(rowHeights).filter(([rowIndex]) => {
      const numericIndex = Number(rowIndex)
      return Number.isInteger(numericIndex) && numericIndex >= 0 && numericIndex < trimmedRowCount
    }),
  )

  const trimmedColumnWidths = Object.fromEntries(
    Object.entries(columnWidths).filter(([columnIndex]) => {
      const numericIndex = Number(columnIndex)
      return Number.isInteger(numericIndex) && numericIndex >= 0 && numericIndex < trimmedColumnCount
    }),
  )

  const trimmedHiddenRows = hiddenRows.filter((range) => trimmedRowCount > 0 && range.start < trimmedRowCount)
    .map((range) => ({ start: range.start, end: Math.min(range.end, trimmedRowCount - 1) }))
  const trimmedHiddenColumns = hiddenColumns.filter((range) => trimmedColumnCount > 0 && range.start < trimmedColumnCount)
    .map((range) => ({ start: range.start, end: Math.min(range.end, trimmedColumnCount - 1) }))
  const trimmedMergedRanges = mergedRanges.filter((rangeA1) => {
    const parsed = parseA1Range(rangeA1)
    return !!parsed && parsed.startRow < trimmedRowCount && parsed.startColumn < trimmedColumnCount
  })

  return {
    ...sheet,
    values: trimmedValues,
    formulas: trimmedFormulas,
    styles: trimmedStyles,
    rowHeights: trimmedRowHeights,
    columnWidths: trimmedColumnWidths,
    hiddenRows: trimmedHiddenRows,
    hiddenColumns: trimmedHiddenColumns,
    mergedRanges: trimmedMergedRanges,
  }
}

function normalizeIndexRanges(value: PersistedIndexRange[] | null | undefined): PersistedIndexRange[] {
  if (!Array.isArray(value)) return []
  return value
    .map((range) => ({
      start: Number(range?.start),
      end: Number(range?.end),
    }))
    .filter((range): range is PersistedIndexRange =>
      Number.isInteger(range.start) &&
      Number.isInteger(range.end) &&
      range.start >= 0 &&
      range.end >= range.start,
    )
}

function normalizeSheetView(value: PersistedSheetView | null | undefined): PersistedSheetView {
  if (!isPlainRecord(value)) return {}
  const sheetView: PersistedSheetView = {}
  const frozenRows = Number(value.frozenRows)
  const frozenColumns = Number(value.frozenColumns)
  if (Number.isInteger(frozenRows) && frozenRows > 0) sheetView.frozenRows = frozenRows
  if (Number.isInteger(frozenColumns) && frozenColumns > 0) sheetView.frozenColumns = frozenColumns
  if (typeof value.hiddenGridlines === 'boolean') sheetView.hiddenGridlines = value.hiddenGridlines
  if (typeof value.gridlinesColor === 'string' && value.gridlinesColor.trim()) {
    sheetView.gridlinesColor = value.gridlinesColor.trim()
  }
  if (typeof value.tabColor === 'string' && value.tabColor.trim()) {
    sheetView.tabColor = value.tabColor.trim()
  }
  return sheetView
}

function normalizeFileGroups(value: PersistedFileGroup[] | null | undefined, sheets: PersistedSheet[]): PersistedFileGroup[] {
  if (Array.isArray(value)) {
    const usedFileIds = new Set(sheets.map((sheet) => sheet.sourceFileId).filter(isNonEmptyString))
    return value
      .filter((group) => group && typeof group.id === 'string' && typeof group.fileName === 'string')
      .map((group) => ({
        id: group.id,
        name: group.name || stripExcelExtension(group.fileName) || group.id,
        fileName: normalizeFileNameInput(group.fileName),
        createdAt: group.createdAt || new Date().toISOString(),
      }))
      .filter((group) => usedFileIds.has(group.id))
  }

  return createGroupsFromSheets(sheets, 'workbook.xlsx')
}

function createGroupsFromSheets(sheets: PersistedSheet[], fallbackFileName: string): PersistedFileGroup[] {
  const groups = new Map<string, PersistedFileGroup>()
  for (const sheet of sheets) {
    const id = sheet.sourceFileId || 'default'
    if (!groups.has(id)) {
      const fileName = normalizeFileNameInput(sheet.sourceFileName || fallbackFileName)
      groups.set(id, {
        id,
        name: stripExcelExtension(fileName) || id,
        fileName,
        createdAt: new Date().toISOString(),
      })
    }
  }
  return [...groups.values()]
}

function createEmptyPersistedWorkbook(): PersistedWorkbook {
  return {
    version: 'grid-v1',
    name: 'Workbook',
    fileName: 'workbook.xlsx',
    activeSheetName: 'Sheet1',
    activeFileId: '',
    fileGroups: [],
    sheets: [],
  }
}

function createEmptyImportSummary(): Required<PersistedImportSummary> {
  return {
    sheets: 0,
    rows: 0,
    columns: 0,
    cells: 0,
    formulas: 0,
    styledCells: 0,
    mergedRanges: 0,
    hiddenRows: 0,
    hiddenColumns: 0,
  }
}

function addImportSummary(target: Required<PersistedImportSummary>, source: PersistedImportSummary | null | undefined): void {
  if (!source) return
  target.sheets += toSafeCount(source.sheets)
  target.rows += toSafeCount(source.rows)
  target.columns += toSafeCount(source.columns)
  target.cells += toSafeCount(source.cells)
  target.formulas += toSafeCount(source.formulas)
  target.styledCells += toSafeCount(source.styledCells)
  target.mergedRanges += toSafeCount(source.mergedRanges)
  target.hiddenRows += toSafeCount(source.hiddenRows)
  target.hiddenColumns += toSafeCount(source.hiddenColumns)
}

function toSafeCount(value: unknown): number {
  return typeof value === 'number' && Number.isFinite(value) && value > 0 ? value : 0
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0
}

function formatInteger(value: number): string {
  return Math.round(value).toLocaleString('zh-CN')
}

function logWorkbookStorageResult(action: string, result: WorkbookStorageResult): void {
  const message = `[univer-agent] workbook ${action}: ${result.ok ? 'saved' : 'failed'} via ${result.storage}; ${formatSnapshotStats(result.stats)}`
  if (result.ok) {
    console.info(message, result.error ? { fallbackReason: result.error } : undefined)
  } else {
    console.warn(message, result.error)
  }
}

function logWorkbookSnapshotStats(action: string, workbook: PersistedWorkbook): void {
  console.info(`[univer-agent] workbook ${action}: ${formatSnapshotStats(getWorkbookSnapshotStats(workbook))}`)
}

function formatSnapshotStats(stats: ReturnType<typeof getWorkbookSnapshotStats>): string {
  return [
    `${formatInteger(stats.sheets)} sheets`,
    `${formatInteger(stats.fileGroups)} file groups`,
    `${formatInteger(stats.styledCells)} styled cells`,
    `${formatInteger(stats.mergedRanges)} merged ranges`,
    formatBytes(stats.bytes),
  ].join(', ')
}

function formatBytes(value: number): string {
  if (!Number.isFinite(value) || value <= 0) return '0 B'
  if (value < 1024) return `${Math.round(value)} B`
  if (value < 1024 * 1024) return `${(value / 1024).toFixed(1)} KB`
  return `${(value / 1024 / 1024).toFixed(2)} MB`
}

function chooseWorkbookForAppend(
  liveWorkbook: PersistedWorkbook | null,
  persistedWorkbook: PersistedWorkbook | null,
): PersistedWorkbook | null {
  const live = liveWorkbook ? normalizePersistedWorkbook(liveWorkbook) : null
  const persisted = persistedWorkbook ? normalizePersistedWorkbook(persistedWorkbook) : null

  if (!persisted) return live
  if (!live) return persisted

  const liveStats = getWorkbookSnapshotStats(live)
  const persistedStats = getWorkbookSnapshotStats(persisted)
  const removedPersistedSheets = getRemovedPersistedSheets(persisted, live)
  const lostAllSheets = liveStats.sheets === 0 && persistedStats.sheets > 0
  const lostExistingStyles = liveStats.sheets === persistedStats.sheets &&
    removedPersistedSheets.length === 0 &&
    persistedStats.styledCells > 0 &&
    liveStats.styledCells < persistedStats.styledCells

  if (lostAllSheets) {
    console.warn('[univer-agent] Live workbook capture looks style-regressed; using persisted snapshot instead.', {
      live: liveStats,
      persisted: persistedStats,
    })
    return persisted
  }

  if (lostExistingStyles) {
    console.warn('[univer-agent] Live workbook capture lost existing styles; restoring presentation data from persisted snapshot.', {
      live: liveStats,
      persisted: persistedStats,
    })
    return mergePersistedPresentationIntoLiveWorkbook(live, persisted)
  }

  return live
}

function getRemovedPersistedSheets(
  persistedWorkbook: PersistedWorkbook,
  liveWorkbook: PersistedWorkbook,
): PersistedSheet[] {
  const liveSheetIds = new Set(
    liveWorkbook.sheets
      .map((sheet) => sheet.sheetId)
      .filter(isNonEmptyString),
  )
  const liveSheetNames = new Set(liveWorkbook.sheets.map((sheet) => sheet.name))

  return persistedWorkbook.sheets.filter((sheet) => {
    if (isNonEmptyString(sheet.sheetId)) {
      return !liveSheetIds.has(sheet.sheetId)
    }

    return !liveSheetNames.has(sheet.name)
  })
}

function mergePersistedPresentationIntoLiveWorkbook(
  liveWorkbook: PersistedWorkbook,
  persistedWorkbook: PersistedWorkbook,
): PersistedWorkbook {
  const persistedSheetsById = new Map(
    persistedWorkbook.sheets
      .filter((sheet) => isNonEmptyString(sheet.sheetId))
      .map((sheet) => [sheet.sheetId as string, sheet]),
  )
  const persistedSheetsByName = new Map(persistedWorkbook.sheets.map((sheet) => [sheet.name, sheet]))

  return {
    ...liveWorkbook,
    fileGroups: liveWorkbook.fileGroups?.length ? liveWorkbook.fileGroups : persistedWorkbook.fileGroups,
    activeFileId: liveWorkbook.activeFileId || persistedWorkbook.activeFileId,
    sheets: liveWorkbook.sheets.map((liveSheet) => {
      const persistedSheet = liveSheet.sheetId
        ? persistedSheetsById.get(liveSheet.sheetId)
        : persistedSheetsByName.get(liveSheet.name)
      if (!persistedSheet) return liveSheet

      return {
        ...liveSheet,
        styles: mergeStyleMatrices(liveSheet.styles, persistedSheet.styles),
        columnWidths: { ...(persistedSheet.columnWidths ?? {}), ...(liveSheet.columnWidths ?? {}) },
        rowHeights: { ...(persistedSheet.rowHeights ?? {}), ...(liveSheet.rowHeights ?? {}) },
        hiddenRows: liveSheet.hiddenRows?.length ? liveSheet.hiddenRows : persistedSheet.hiddenRows,
        hiddenColumns: liveSheet.hiddenColumns?.length ? liveSheet.hiddenColumns : persistedSheet.hiddenColumns,
        mergedRanges: liveSheet.mergedRanges?.length ? liveSheet.mergedRanges : persistedSheet.mergedRanges,
        sheetView: { ...(persistedSheet.sheetView ?? {}), ...(liveSheet.sheetView ?? {}) },
      }
    }),
  }
}

function mergeStyleMatrices(
  liveStyles: Array<Array<Record<string, unknown> | null>> | undefined,
  persistedStyles: Array<Array<Record<string, unknown> | null>> | undefined,
): Array<Array<Record<string, unknown> | null>> {
  const liveMatrix = normalizeMatrix<Record<string, unknown> | null>(liveStyles)
  const persistedMatrix = normalizeMatrix<Record<string, unknown> | null>(persistedStyles)
  const rowCount = Math.max(liveMatrix.length, persistedMatrix.length)
  const merged: Array<Array<Record<string, unknown> | null>> = []

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const liveRow = liveMatrix[rowIndex] ?? []
    const persistedRow = persistedMatrix[rowIndex] ?? []
    const columnCount = Math.max(liveRow.length, persistedRow.length)
    const mergedRow: Array<Record<string, unknown> | null> = []

    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      mergedRow[columnIndex] = mergeCellStyles(
        liveRow[columnIndex] ?? null,
        persistedRow[columnIndex] ?? null,
      )
    }

    merged[rowIndex] = mergedRow
  }

  return merged
}

function mergeCellStyles(
  liveStyle: Record<string, unknown> | null,
  persistedStyle: Record<string, unknown> | null,
): Record<string, unknown> | null {
  if (!liveStyle) {
    return persistedStyle ? { ...persistedStyle } : null
  }

  if (!persistedStyle) {
    return { ...liveStyle }
  }

  const mergedStyle: Record<string, unknown> = {
    ...persistedStyle,
    ...liveStyle,
  }
  const mergedBorder = mergeCellBorders(liveStyle.border, persistedStyle.border)

  if (mergedBorder) {
    mergedStyle.border = mergedBorder
  }

  return mergedStyle
}

function mergeCellBorders(
  liveBorder: unknown,
  persistedBorder: unknown,
): Record<string, unknown> | null {
  const liveRecord = isPlainRecord(liveBorder) ? liveBorder : null
  const persistedRecord = isPlainRecord(persistedBorder) ? persistedBorder : null

  if (!liveRecord) {
    return persistedRecord ? { ...persistedRecord } : null
  }

  if (!persistedRecord) {
    return { ...liveRecord }
  }

  const mergedBorder: Record<string, unknown> = {
    ...persistedRecord,
    ...liveRecord,
  }
  const sides = ['top', 'right', 'bottom', 'left']

  sides.forEach((side) => {
    const liveEdge = liveRecord[side]
    const persistedEdge = persistedRecord[side]

    if (isPlainRecord(liveEdge) && isPlainRecord(persistedEdge)) {
      mergedBorder[side] = {
        ...persistedEdge,
        ...liveEdge,
      }
      return
    }

    if (liveEdge === undefined && persistedEdge !== undefined) {
      mergedBorder[side] = persistedEdge
    }
  })

  return Object.keys(mergedBorder).length > 0 ? mergedBorder : null
}

function createFileGroup(fileName: string): PersistedFileGroup {
  const normalizedFileName = normalizeFileNameInput(fileName)
  const name = stripExcelExtension(normalizedFileName) || 'Workbook'
  return {
    id: `${sanitizeSheetName(name).toLowerCase()}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`,
    name,
    fileName: normalizedFileName,
    createdAt: new Date().toISOString(),
  }
}

interface FileOperationTarget {
  fileId: string
  group: PersistedFileGroup
  sheets: PersistedSheet[]
}

function resolveFileOperationTarget(workbook: PersistedWorkbook, fallbackFileId = ''): FileOperationTarget | null {
  const normalizedWorkbook = normalizePersistedWorkbook(workbook)
  const groups = normalizedWorkbook.fileGroups ?? []
  if (groups.length === 0) return null

  const activeSheet = normalizedWorkbook.sheets.find((sheet) => sheet.name === normalizedWorkbook.activeSheetName)
  const activeSheetFileId = activeSheet?.sourceFileId ?? ''
  const fileId = groups.some((group) => group.id === activeSheetFileId)
    ? activeSheetFileId
    : (groups.some((group) => group.id === fallbackFileId) ? fallbackFileId : groups[0].id)
  const group = groups.find((item) => item.id === fileId)
  if (!group) return null

  const sheets = normalizedWorkbook.sheets.filter((sheet) => sheet.sourceFileId === fileId)
  return sheets.length > 0 ? { fileId, group, sheets } : null
}

function getFileGroupSheetCounts(workbook: PersistedWorkbook | null): Map<string, number> {
  const counts = new Map<string, number>()
  for (const sheet of workbook?.sheets ?? []) {
    if (!sheet.sourceFileId) continue
    counts.set(sheet.sourceFileId, (counts.get(sheet.sourceFileId) ?? 0) + 1)
  }
  return counts
}

function appendImportedWorkbook(
  targetWorkbook: PersistedWorkbook,
  importedWorkbook: PersistedWorkbook,
  group: PersistedFileGroup,
  usedSheetNames: Set<string>,
): PersistedWorkbook {
  const importedSheets = importedWorkbook.sheets.map((sheet) => {
    const originalSheetName = sheet.originalSheetName || sheet.name
    const displayName = getUniqueSheetName(`${group.name}_${originalSheetName}`, usedSheetNames)
    usedSheetNames.add(displayName)
    return {
      ...sheet,
      name: displayName,
      sourceFileId: group.id,
      sourceFileName: group.fileName,
      originalSheetName,
    }
  })

  return {
    ...targetWorkbook,
    name: 'Workbook',
    fileName: group.fileName,
    activeFileId: group.id,
    fileGroups: [...(targetWorkbook.fileGroups ?? []), group],
    activeSheetName: importedSheets[0]?.name ?? targetWorkbook.activeSheetName,
    sheets: [...targetWorkbook.sheets, ...importedSheets],
  }
}

function removeFileGroupFromWorkbook(workbook: PersistedWorkbook, fileId: string): PersistedWorkbook {
  const remainingGroups = (workbook.fileGroups ?? []).filter((group) => group.id !== fileId)
  const remainingSheets = workbook.sheets.filter((sheet) => sheet.sourceFileId !== fileId)
  const nextActiveFileId = remainingGroups[0]?.id ?? ''
  const nextActiveGroup = remainingGroups.find((group) => group.id === nextActiveFileId)
  const nextActiveSheet = remainingSheets.find((sheet) => sheet.sourceFileId === nextActiveFileId) ?? remainingSheets[0]

  return normalizePersistedWorkbook({
    ...workbook,
    name: nextActiveGroup?.name ?? 'Workbook',
    fileName: nextActiveGroup?.fileName ?? 'workbook.xlsx',
    activeFileId: nextActiveFileId,
    fileGroups: remainingGroups,
    activeSheetName: nextActiveSheet?.name ?? 'Sheet1',
    sheets: remainingSheets,
  })
}

function getWorkbookForFileGroup(workbook: PersistedWorkbook, fileId: string): PersistedWorkbook {
  if (!fileId) return workbook

  const group = workbook.fileGroups?.find((item) => item.id === fileId)
  const usedExportNames = new Set<string>()
  const sheets = workbook.sheets
    .filter((sheet) => sheet.sourceFileId === fileId)
    .map((sheet) => {
      const exportName = getUniqueSheetName(getExportSheetName(sheet.originalSheetName || sheet.name), usedExportNames)
      usedExportNames.add(exportName)
      return {
        ...sheet,
        name: exportName,
      }
    })

  return {
    ...workbook,
    name: group?.name ?? workbook.name,
    fileName: group?.fileName ?? workbook.fileName,
    activeFileId: fileId,
    fileGroups: group ? [group] : [],
    activeSheetName: sheets[0]?.name ?? workbook.activeSheetName,
    sheets,
  }
}

function syncSheetMetadataFromWorkbook(
  workbook: LocalWorkbook,
  persisted: PersistedWorkbook,
  sheetMeta: Map<string, SheetSourceMeta>,
): void {
  sheetMeta.clear()
  for (const sheetData of persisted.sheets) {
    const sheet = workbook.getSheetByName(sheetData.name)
    if (!sheet) continue
    const meta = {
      sourceFileId: sheetData.sourceFileId || persisted.activeFileId || 'default',
      sourceFileName: sheetData.sourceFileName || persisted.fileName || 'workbook.xlsx',
      originalSheetName: sheetData.originalSheetName || sheetData.name,
    }
    sheetMeta.set(sheet.getSheetId(), meta)
    sheetMeta.set(sheet.getSheetName(), meta)
  }
}

function getSheetSourceMeta(
  sheet: LocalSheet,
  sheetMeta: Map<string, SheetSourceMeta>,
  persistedWorkbook: PersistedWorkbook | null,
): SheetSourceMeta | null {
  const existingMeta = sheetMeta.get(sheet.getSheetId()) ?? sheetMeta.get(sheet.getSheetName())
  if (existingMeta) return existingMeta

  const persistedSheet = persistedWorkbook?.sheets.find((item) =>
    item.sheetId === sheet.getSheetId() || item.name === sheet.getSheetName(),
  )
  if (!persistedSheet?.sourceFileId) return null

  return {
    sourceFileId: persistedSheet.sourceFileId,
    sourceFileName: persistedSheet.sourceFileName || persistedWorkbook?.fileName || 'workbook.xlsx',
    originalSheetName: persistedSheet.originalSheetName || persistedSheet.name,
  }
}

function getUniqueSheetName(baseName: string, usedNames: Set<string>): string {
  const base = sanitizeSheetName(baseName) || 'Sheet'
  let candidate = base
  let index = 2

  while (usedNames.has(candidate)) {
    const suffix = `_${index}`
    candidate = `${base.slice(0, Math.max(1, 31 - suffix.length))}${suffix}`
    index += 1
  }

  return candidate
}

function getExportSheetName(name: string): string {
  return sanitizeSheetName(name) || 'Sheet'
}

function sanitizeSheetName(name: string): string {
  const cleaned = name
    .split('')
    .map((character) => ('[]:*?/\\'.includes(character) ? '_' : character))
    .join('')
    .trim()
  return cleaned.slice(0, 31)
}

function capturePersistedSheet(sheet: LocalSheet, sheetMeta: Map<string, SheetSourceMeta>): PersistedSheet {
  const lastRow = sheet.getLastRow()
  const lastColumn = sheet.getLastColumn()
  const hasData = lastRow >= 0 && lastColumn >= 0
  const columnWidths: Record<string, number> = {}
  const rowHeights: Record<string, number> = {}
  const mergedRanges = getMergedRangeA1List(sheet)
  const hiddenRows = getHiddenRowRanges(sheet)
  const hiddenColumns = getHiddenColumnRanges(sheet)
  const sheetView = getPersistedSheetView(sheet)
  const meta = sheetMeta.get(sheet.getSheetId()) ?? sheetMeta.get(sheet.getSheetName()) ?? null
  const baseSheet = {
    name: sheet.getSheetName(),
    sheetId: sheet.getSheetId(),
    sourceFileId: meta?.sourceFileId,
    sourceFileName: meta?.sourceFileName,
    originalSheetName: meta?.originalSheetName ?? sheet.getSheetName(),
    hiddenRows,
    hiddenColumns,
    sheetView,
  }

  for (let columnIndex = 0; columnIndex <= Math.max(lastColumn, 0); columnIndex += 1) {
    const width = safeNumber(() => sheet.getColumnWidth(columnIndex))
    if (width !== null) columnWidths[String(columnIndex)] = width
  }

  for (let rowIndex = 0; rowIndex <= Math.max(lastRow, 0); rowIndex += 1) {
    const height = safeNumber(() => sheet.getRowHeight(rowIndex))
    if (height !== null) rowHeights[String(rowIndex)] = height
  }

  if (!hasData) {
    return trimPersistedSheetBounds({
      ...baseSheet,
      values: [],
      formulas: [],
      styles: [],
      columnWidths,
      rowHeights,
      mergedRanges,
    })
  }

  const range = sheet.getRange(0, 0, lastRow + 1, lastColumn + 1)
  const rawValues = normalizeMatrix<unknown>(range.getValues())
  const rawFormulas = normalizeMatrix<string>(range.getFormulas())
  const rawCellData = normalizeMatrix<Record<string, unknown> | null>(range.getCellDataGrid?.() ?? [])
  const rawNumberFormats = normalizeMatrix<string>(range.getNumberFormats?.() ?? [])
  const styles = rawCellData.map((row, rowIndex) =>
    row.map((cellData, columnIndex) => mergePersistedStyle(
      extractPersistedStyle(cellData),
      rawNumberFormats[rowIndex]?.[columnIndex],
    )),
  )

  return trimPersistedSheetBounds({
    ...baseSheet,
    values: rawValues.map((row) => row.map((value) => normalizeCellValue(value))),
    formulas: rawFormulas.map((row) => row.map((formula) => formula ?? '')),
    styles,
    columnWidths,
    rowHeights,
    mergedRanges,
  })
}

function applyPersistedWorkbookToWorkbook(
  workbook: LocalWorkbook,
  persisted: PersistedWorkbook,
  univerAPI: { Enum?: Record<string, Record<string, unknown>> },
): void {
  const targetNames = new Set(persisted.sheets.map((sheet) => sheet.name))
  const currentSheets = workbook.getSheets()

  currentSheets.forEach((sheet) => {
    if (!targetNames.has(sheet.getSheetName()) && workbook.getSheets().length > 1) {
      try {
        workbook.deleteSheet(sheet)
      } catch {
        workbook.deleteSheet(sheet.getSheetId())
      }
    }
  })

  const firstSheet = workbook.getActiveSheet()
  persisted.sheets.forEach((sheetData, index) => {
    let targetSheet = workbook.getSheetByName(sheetData.name)
    if (!targetSheet && index === 0 && firstSheet) {
      targetSheet = firstSheet
      if (targetSheet.getSheetName() !== sheetData.name) {
        targetSheet.setName(sheetData.name)
      }
    }

    if (!targetSheet) {
      targetSheet = workbook.insertSheet(sheetData.name)
    }

    clearSheet(targetSheet)
    applyPersistedSheet(targetSheet, sheetData, univerAPI)
  })
}

function applyPersistedSheet(
  sheet: LocalSheet,
  sheetData: PersistedSheet,
  univerAPI: { Enum?: Record<string, Record<string, unknown>> },
): void {
  const values = normalizeMatrix<CellValue>(sheetData.values)
  const formulas = normalizeMatrix<string>(sheetData.formulas)
  const styles = normalizeMatrix<Record<string, unknown> | null>(sheetData.styles)
  const columnWidths = isPlainRecord(sheetData.columnWidths) ? sheetData.columnWidths : {}
  const rowHeights = isPlainRecord(sheetData.rowHeights) ? sheetData.rowHeights : {}
  const mergedRanges = Array.isArray(sheetData.mergedRanges) ? sheetData.mergedRanges : []
  const hiddenRows = normalizeIndexRanges(sheetData.hiddenRows)
  const hiddenColumns = normalizeIndexRanges(sheetData.hiddenColumns)
  const sheetView = normalizeSheetView(sheetData.sheetView)
  const rowCount = Math.max(
    values.length,
    formulas.length,
    styles.length,
    getMaxIndexFromRecord(rowHeights as Record<string, number>) + 1,
    getMaxCoveredIndexFromRanges(hiddenRows) + 1,
    getMaxCoveredIndexFromMergedRanges(mergedRanges, 'row') + 1,
    0,
  )
  const columnCount = Math.max(
    0,
    ...values.map((row) => row.length),
    ...formulas.map((row) => row.length),
    ...styles.map((row) => row.length),
    getMaxIndexFromRecord(columnWidths as Record<string, number>) + 1,
    getMaxCoveredIndexFromRanges(hiddenColumns) + 1,
    getMaxCoveredIndexFromMergedRanges(mergedRanges, 'column') + 1,
  )

  const visibleRowCount = Math.max(MIN_VISIBLE_ROW_COUNT, rowCount)
  const visibleColumnCount = Math.max(MIN_VISIBLE_COLUMN_COUNT, columnCount)

  ensureSheetCapacity(sheet, visibleRowCount, visibleColumnCount)

  if (rowCount > 0 && columnCount > 0) {
    sheet.getRange(0, 0, rowCount, columnCount).setValues(
      createCellDataMatrix(values, styles, rowCount, columnCount),
    )

    formulas.forEach((row, rowIndex) => {
      row.forEach((formula, columnIndex) => {
        if (formula) {
          sheet.getRange(rowIndex, columnIndex).setFormula(formula)
        }
      })
    })

    styles.forEach((row, rowIndex) => {
      row.forEach((style, columnIndex) => {
        if (style) {
          applyPersistedStyleToRange(sheet.getRange(rowIndex, columnIndex), style, univerAPI)
        }
      })
    })
  }

  Object.entries(columnWidths).forEach(([columnIndex, width]) => {
    const index = Number(columnIndex)
    if (Number.isInteger(index) && Number(width) > 0) {
      sheet.setColumnWidth(index, Number(width))
    }
  })

  Object.entries(rowHeights).forEach(([rowIndex, height]) => {
    const index = Number(rowIndex)
    if (Number.isInteger(index) && Number(height) > 0) {
      sheet.setRowHeight(index, Number(height))
    }
  })

  hiddenRows.forEach((range) => {
    sheet.hideRows?.(range.start, range.end - range.start + 1)
  })

  hiddenColumns.forEach((range) => {
    sheet.hideColumns?.(range.start, range.end - range.start + 1)
  })

  applyPersistedSheetView(sheet, sheetView)

  mergedRanges.forEach((rangeA1) => {
    try {
      sheet.getRange(rangeA1).merge()
    } catch {
      // Invalid or unsupported merged ranges should not block workbook import.
    }
  })
}

function ensureSheetCapacity(sheet: LocalSheet, requiredRowCount: number, requiredColumnCount: number): void {
  const coreSheet = sheet.getSheet?.()
  const currentRowCount = safeInteger(() =>
    Number(coreSheet?.getRowCount?.() ?? sheet.getLastRow() + 1),
  )
  const currentColumnCount = safeInteger(() =>
    Number(coreSheet?.getColumnCount?.() ?? sheet.getLastColumn() + 1),
  )

  if (requiredRowCount > currentRowCount) {
    try {
      sheet.setRowCount?.(requiredRowCount)
    } catch {
      try {
        coreSheet?.setRowCount?.(requiredRowCount)
      } catch {
        // Best effort. If capacity cannot be expanded, the following write will still surface the real error.
      }
    }
  }

  if (requiredColumnCount > currentColumnCount) {
    try {
      sheet.setColumnCount?.(requiredColumnCount)
    } catch {
      try {
        coreSheet?.setColumnCount?.(requiredColumnCount)
      } catch {
        // Best effort. If capacity cannot be expanded, the following write will still surface the real error.
      }
    }
  }
}

function createCellDataMatrix(
  values: CellValue[][],
  styles: Array<Array<Record<string, unknown> | null>>,
  rowCount: number,
  columnCount: number,
): UniverCellData[][] {
  return Array.from({ length: rowCount }, (_, rowIndex) =>
    Array.from({ length: columnCount }, (_, columnIndex) => {
      const value = values[rowIndex]?.[columnIndex]
      const cellData: UniverCellData = {
        v: value === undefined ? null : value,
      }
      const univerStyle = persistedStyleToUniverCellStyle(styles[rowIndex]?.[columnIndex] ?? null)
      if (univerStyle) cellData.s = univerStyle
      return cellData
    }),
  )
}

function persistedStyleToUniverCellStyle(style: Record<string, unknown> | null): Record<string, unknown> | null {
  if (!style) return null
  const result: Record<string, unknown> = {}

  if (typeof style.fontFamily === 'string') result.ff = style.fontFamily
  if (style.fontWeight === 'bold') result.bl = 1
  if (style.fontStyle === 'italic') result.it = 1
  if (style.underline === true || style.fontLine === 'underline') result.ul = { s: 1 }
  if (style.strikethrough === true || style.fontLine === 'line-through') result.st = { s: 1 }
  if (typeof style.fontSize === 'number') result.fs = style.fontSize
  if (typeof style.fontColor === 'string') result.cl = { rgb: style.fontColor }
  if (typeof style.background === 'string') result.bg = { rgb: style.background }
  if (typeof style.horizontalAlignment === 'string') result.ht = persistedHorizontalAlignmentToUniver(style.horizontalAlignment)
  if (typeof style.verticalAlignment === 'string') result.vt = persistedVerticalAlignmentToUniver(style.verticalAlignment)
  if (typeof style.textRotation === 'number') result.tr = { a: style.textRotation }
  if (style.wrap) result.tb = 3
  if (typeof style.numFmt === 'string' && style.numFmt.trim()) {
    result.n = { pattern: style.numFmt }
  }
  const univerBorder = persistedBorderToUniverBorder(style.border)
  if (univerBorder) result.bd = univerBorder

  return Object.keys(result).length > 0 ? result : null
}

function applyPersistedStyleToRange(
  range: LocalRange,
  style: Record<string, unknown>,
  univerAPI: { Enum?: Record<string, Record<string, unknown>> },
): void {
  if (typeof style.fontFamily === 'string') range.setFontFamily?.(style.fontFamily)
  if (style.fontWeight === 'bold') range.setFontWeight('bold')
  if (style.fontStyle === 'italic') range.setFontStyle('italic')
  if (style.underline === true || style.fontLine === 'underline') range.setFontLine?.('underline')
  if (style.strikethrough === true || style.fontLine === 'line-through') range.setFontLine?.('line-through')
  if (typeof style.fontSize === 'number') range.setFontSize(style.fontSize)
  if (typeof style.fontColor === 'string') range.setFontColor(style.fontColor)
  if (typeof style.background === 'string') range.setBackground(style.background)
  if (typeof style.horizontalAlignment === 'string') range.setHorizontalAlignment(normalizeHorizontalAlignment(style.horizontalAlignment))
  if (typeof style.verticalAlignment === 'string') range.setVerticalAlignment?.(normalizeVerticalAlignment(style.verticalAlignment))
  if (typeof style.textRotation === 'number') range.setTextRotation?.(style.textRotation)
  if (style.wrap) {
    const wrapStrategy = univerAPI.Enum?.WrapStrategy?.WRAP
    if (wrapStrategy) range.setWrapStrategy(wrapStrategy)
  }
  if (typeof style.numFmt === 'string' && style.numFmt.trim() && range.setNumberFormat) {
    range.setNumberFormat(style.numFmt)
  }
  if (style.border) {
    applyPersistedBorderToRange(range, style.border, univerAPI)
  }
}

function clearSheet(sheet: LocalSheet): void {
  const lastRow = sheet.getLastRow()
  const lastColumn = sheet.getLastColumn()

  getHiddenRowRanges(sheet).forEach((range) => {
    try {
      sheet.unhideRow?.(sheet.getRange(`${range.start + 1}:${range.end + 1}`))
    } catch {
      // Best-effort cleanup before replacing workbook content.
    }
  })

  getHiddenColumnRanges(sheet).forEach((range) => {
    try {
      sheet.unhideColumn?.(sheet.getRange(`${columnIndexToLetters(range.start)}:${columnIndexToLetters(range.end)}`))
    } catch {
      // Best-effort cleanup before replacing workbook content.
    }
  })

  try {
    sheet.getMergedRanges().forEach((range) => range.breakApart())
  } catch {
    // Best-effort cleanup before replacing workbook content.
  }

  if (lastRow >= 0 && lastColumn >= 0) {
    sheet.getRange(0, 0, lastRow + 1, lastColumn + 1).clear()
  }
}

function applyPersistedSheetView(sheet: LocalSheet, sheetView: PersistedSheetView): void {
  sheet.setFrozenRows?.(sheetView.frozenRows ?? 0)
  sheet.setFrozenColumns?.(sheetView.frozenColumns ?? 0)
  sheet.setHiddenGridlines?.(sheetView.hiddenGridlines ?? false)
  if (sheetView.tabColor) sheet.setTabColor?.(sheetView.tabColor)
}

function applyPersistedBorderToRange(
  range: LocalRange,
  border: unknown,
  univerAPI: { Enum?: Record<string, Record<string, unknown>> },
): void {
  if (!isPlainRecord(border)) return

  const borderTypeEnum = univerAPI.Enum?.BorderType
  const borderStyleEnum = univerAPI.Enum?.BorderStyleTypes
  if (!borderTypeEnum || !borderStyleEnum) return

  const sides: Array<[string, string]> = [
    ['top', 'TOP'],
    ['right', 'RIGHT'],
    ['bottom', 'BOTTOM'],
    ['left', 'LEFT'],
  ]

  sides.forEach(([side, typeKey]) => {
    const edge = border[side]
    if (!isPlainRecord(edge)) return

    const borderType = borderTypeEnum[typeKey]
    const borderStyle = getUniverBorderStyle(edge.style, borderStyleEnum)
    if (!borderType || !borderStyle) return

    range.setBorder(borderType, borderStyle, typeof edge.color === 'string' ? edge.color : '#000000')
  })
}

function persistedBorderToUniverBorder(value: unknown): Record<string, unknown> | null {
  if (!isPlainRecord(value)) return null

  const sides: Array<[string, string]> = [
    ['top', 't'],
    ['right', 'r'],
    ['bottom', 'b'],
    ['left', 'l'],
  ]
  const result: Record<string, unknown> = {}

  sides.forEach(([side, key]) => {
    const edge = value[side]
    if (!isPlainRecord(edge)) return

    result[key] = {
      s: edge.style,
      cl: typeof edge.color === 'string' ? { rgb: edge.color } : { rgb: '#000000' },
    }
  })

  return Object.keys(result).length > 0 ? result : null
}

function getUniverBorderStyle(style: unknown, borderStyleEnum: Record<string, unknown>): unknown {
  if (typeof style === 'number') return style
  const key = excelBorderStyleToUniverKey(style)
  return borderStyleEnum[key] ?? borderStyleEnum.THIN
}

function excelBorderStyleToUniverKey(style: unknown): string {
  switch (String(style || '').toLowerCase()) {
    case 'hair':
      return 'HAIR'
    case 'dotted':
      return 'DOTTED'
    case 'dashdot':
      return 'DASH_DOT'
    case 'dashdotdot':
      return 'DASH_DOT_DOT'
    case 'dashed':
      return 'DASHED'
    case 'double':
      return 'DOUBLE'
    case 'medium':
      return 'MEDIUM'
    case 'mediumdashed':
      return 'MEDIUM_DASHED'
    case 'mediumdashdot':
      return 'MEDIUM_DASH_DOT'
    case 'mediumdashdotdot':
      return 'MEDIUM_DASH_DOT_DOT'
    case 'slantdashdot':
      return 'SLANT_DASH_DOT'
    case 'thick':
      return 'THICK'
    case 'thin':
    default:
      return 'THIN'
  }
}

function normalizeHorizontalAlignment(value: string): string {
  if (value === 'centerContinuous') return 'center'
  if (value === 'distributed' || value === 'justify') return 'center'
  if (value === 'fill') return 'left'
  return value
}

function normalizeVerticalAlignment(value: string): string {
  if (value === 'center') return 'middle'
  if (value === 'justify' || value === 'distributed') return 'middle'
  return value
}

function persistedHorizontalAlignmentToUniver(value: string): number | string {
  switch (normalizeHorizontalAlignment(value)) {
    case 'left':
      return 1
    case 'center':
      return 2
    case 'right':
      return 3
    default:
      return value
  }
}

function persistedVerticalAlignmentToUniver(value: string): number | string {
  switch (normalizeVerticalAlignment(value)) {
    case 'top':
      return 1
    case 'middle':
      return 2
    case 'bottom':
      return 3
    default:
      return value
  }
}

function getPersistedSheetView(sheet: LocalSheet): PersistedSheetView {
  const coreSheet = sheet.getSheet?.()
  const frozenRows = safeInteger(() => sheet.getFrozenRows?.() ?? Number(coreSheet?.getFreeze?.()?.ySplit ?? 0))
  const frozenColumns = safeInteger(() => sheet.getFrozenColumns?.() ?? Number(coreSheet?.getFreeze?.()?.xSplit ?? 0))
  const hiddenGridlines = safeBoolean(() =>
    sheet.hasHiddenGridLines?.() ??
    sheet.hasHiddenGridlines?.() ??
    coreSheet?.hasHiddenGridlines?.() ??
    false,
  )
  const gridlinesColor = sheet.getGridLinesColor?.() ?? coreSheet?.getGridlinesColor?.()
  const tabColor = sheet.getTabColor?.() ?? coreSheet?.getTabColor?.()
  const sheetView: PersistedSheetView = {}

  if (frozenRows > 0) sheetView.frozenRows = frozenRows
  if (frozenColumns > 0) sheetView.frozenColumns = frozenColumns
  if (hiddenGridlines) sheetView.hiddenGridlines = true
  if (gridlinesColor) sheetView.gridlinesColor = gridlinesColor
  if (tabColor) sheetView.tabColor = tabColor

  return sheetView
}

function getHiddenRowRanges(sheet: LocalSheet): PersistedIndexRange[] {
  const coreSheet = sheet.getSheet?.()
  return convertCoreRangesToIndexRanges(coreSheet?.getHiddenRows?.(), 'row')
}

function getHiddenColumnRanges(sheet: LocalSheet): PersistedIndexRange[] {
  const coreSheet = sheet.getSheet?.()
  return convertCoreRangesToIndexRanges(coreSheet?.getHiddenCols?.(), 'column')
}

function convertCoreRangesToIndexRanges(
  ranges: Array<Record<string, unknown>> | undefined,
  axis: 'row' | 'column',
): PersistedIndexRange[] {
  if (!Array.isArray(ranges)) return []
  const startKey = axis === 'row' ? 'startRow' : 'startColumn'
  const endKey = axis === 'row' ? 'endRow' : 'endColumn'

  return ranges
    .map((range) => ({
      start: Number(range[startKey]),
      end: Number(range[endKey]),
    }))
    .filter((range): range is PersistedIndexRange =>
      Number.isInteger(range.start) &&
      Number.isInteger(range.end) &&
      range.start >= 0 &&
      range.end >= range.start,
    )
}

function mergePersistedStyle(
  style: Record<string, unknown> | null,
  numberFormat: string | null | undefined,
): Record<string, unknown> | null {
  const normalizedNumberFormat = typeof numberFormat === 'string' ? numberFormat.trim() : ''
  if (!normalizedNumberFormat) return style
  return {
    ...(style ?? {}),
    numFmt: normalizedNumberFormat,
  }
}

function isTextDecorationEnabled(value: unknown): boolean {
  return isPlainRecord(value) && (value.s === 1 || value.s === true)
}

function colorStyleToHex(value: unknown): string | null {
  if (typeof value === 'string') return normalizeCssHexColor(value)
  if (!isPlainRecord(value)) return null
  return typeof value.rgb === 'string' ? normalizeCssHexColor(value.rgb) : null
}

function normalizeCssHexColor(value: unknown): string | null {
  if (typeof value !== 'string') return null
  const trimmed = value.trim()
  const namedColors: Record<string, string> = {
    black: '#000000',
    white: '#FFFFFF',
    red: '#FF0000',
    green: '#008000',
    blue: '#0000FF',
    yellow: '#FFFF00',
    orange: '#FFA500',
    purple: '#800080',
    gray: '#808080',
    grey: '#808080',
    cyan: '#00FFFF',
    magenta: '#FF00FF',
  }
  const named = namedColors[trimmed.toLowerCase()]
  if (named) return named

  const rgbMatch = /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})/i.exec(trimmed)
  if (rgbMatch) {
    const channels = rgbMatch.slice(1, 4).map((channel) => Math.max(0, Math.min(255, Number(channel))))
    return `#${channels.map((channel) => channel.toString(16).padStart(2, '0')).join('').toUpperCase()}`
  }

  const normalized = trimmed.replace(/^#/, '')
  if (/^[0-9a-f]{8}$/i.test(normalized)) return `#${normalized.slice(2).toUpperCase()}`
  if (/^[0-9a-f]{6}$/i.test(normalized)) return `#${normalized.toUpperCase()}`
  if (/^[0-9a-f]{3}$/i.test(normalized)) {
    return `#${normalized.split('').map((part) => part + part).join('').toUpperCase()}`
  }
  return trimmed || null
}

function univerHorizontalAlignmentToPersisted(value: unknown): string | null {
  switch (value) {
    case 1:
      return 'left'
    case 2:
      return 'center'
    case 3:
      return 'right'
    default:
      return typeof value === 'string' ? value : null
  }
}

function univerVerticalAlignmentToPersisted(value: unknown): string | null {
  switch (value) {
    case 1:
      return 'top'
    case 2:
      return 'middle'
    case 3:
      return 'bottom'
    default:
      return typeof value === 'string' ? value : null
  }
}

function getTextRotation(value: unknown): number | null {
  if (typeof value === 'number') return value
  if (!isPlainRecord(value)) return null
  return typeof value.a === 'number' ? value.a : null
}

function univerBorderToPersistedBorder(value: unknown): Record<string, unknown> | null {
  if (!isPlainRecord(value)) return null
  const sides: Array<[string, string]> = [
    ['top', 't'],
    ['right', 'r'],
    ['bottom', 'b'],
    ['left', 'l'],
  ]
  const result: Record<string, unknown> = {}

  sides.forEach(([side, key]) => {
    const edge = value[key]
    if (!isPlainRecord(edge)) return
    result[side] = {
      style: edge.s,
      color: colorStyleToHex(edge.cl) ?? '#000000',
    }
  })

  return Object.keys(result).length > 0 ? result : null
}

function extractPersistedStyle(cellData: Record<string, unknown> | null): Record<string, unknown> | null {
  if (!cellData) return null
  const directStyle = isPlainRecord(cellData.s) ? cellData.s : cellData
  const style: Record<string, unknown> = {}

  if (typeof directStyle.ff === 'string') style.fontFamily = directStyle.ff
  if (directStyle.bl === 1 || directStyle.bold === true || directStyle.fontWeight === 'bold') style.fontWeight = 'bold'
  if (directStyle.it === 1 || directStyle.italic === true || directStyle.fontStyle === 'italic') style.fontStyle = 'italic'
  if (isTextDecorationEnabled(directStyle.ul) || directStyle.underline === true) style.underline = true
  if (isTextDecorationEnabled(directStyle.st) || directStyle.strikethrough === true) style.strikethrough = true
  if (typeof directStyle.fontLine === 'string') style.fontLine = directStyle.fontLine
  if (typeof directStyle.fs === 'number') style.fontSize = directStyle.fs
  if (typeof directStyle.fontSize === 'number') style.fontSize = directStyle.fontSize
  const directFontColor = normalizeCssHexColor(directStyle.fc) ?? normalizeCssHexColor(directStyle.fontColor)
  if (directFontColor) style.fontColor = directFontColor
  const foreground = colorStyleToHex(directStyle.cl)
  if (foreground) style.fontColor = foreground
  const directBackground = normalizeCssHexColor(directStyle.bg) ?? normalizeCssHexColor(directStyle.background)
  if (directBackground) style.background = directBackground
  const background = colorStyleToHex(directStyle.bg)
  if (background) style.background = background
  const horizontalAlignment = univerHorizontalAlignmentToPersisted(directStyle.ht)
  if (horizontalAlignment) style.horizontalAlignment = horizontalAlignment
  if (typeof directStyle.horizontalAlignment === 'string') style.horizontalAlignment = directStyle.horizontalAlignment
  const verticalAlignment = univerVerticalAlignmentToPersisted(directStyle.vt)
  if (verticalAlignment) style.verticalAlignment = verticalAlignment
  if (typeof directStyle.verticalAlignment === 'string') style.verticalAlignment = directStyle.verticalAlignment
  if (directStyle.tb === 3 || directStyle.wrap === true) style.wrap = true
  const textRotation = getTextRotation(directStyle.tr)
  if (textRotation !== null) style.textRotation = textRotation
  const border = directStyle.border ?? univerBorderToPersistedBorder(directStyle.bd)
  if (border) style.border = border
  if (typeof directStyle.numFmt === 'string') style.numFmt = directStyle.numFmt
  if (isPlainRecord(directStyle.n) && typeof directStyle.n.pattern === 'string') {
    style.numFmt = directStyle.n.pattern
  }
  if (isPlainRecord(directStyle.numberFormat) && typeof directStyle.numberFormat.pattern === 'string') {
    style.numFmt = directStyle.numberFormat.pattern
  }

  return Object.keys(style).length > 0 ? style : null
}

function getMergedRangeA1List(sheet: LocalSheet): string[] {
  try {
    return sheet.getMergedRanges().map((range) => range.getA1Notation?.()).filter((range): range is string => !!range)
  } catch {
    return []
  }
}

function normalizeCellValue(value: unknown): CellValue {
  if (value === null || value === undefined || value === '') return null
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value
  return String(value)
}

function columnIndexToLetters(index: number): string {
  let value = index + 1
  let letters = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    letters = String.fromCharCode(65 + remainder) + letters
    value = Math.floor((value - 1) / 26)
  }
  return letters || 'A'
}

function safeNumber(reader: () => number): number | null {
  try {
    const value = reader()
    return Number.isFinite(value) ? value : null
  } catch {
    return null
  }
}

function safeInteger(reader: () => number): number {
  try {
    const value = reader()
    return Number.isInteger(value) && value > 0 ? value : 0
  } catch {
    return 0
  }
}

function safeBoolean(reader: () => boolean): boolean {
  try {
    return reader() === true
  } catch {
    return false
  }
}

function isPlainRecord(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value)
}

function stripExcelExtension(fileName: string): string {
  return fileName.replace(/\.(xlsx|xlsm|xls)$/i, '')
}

function normalizeFileNameInput(value: string): string {
  const cleaned = value
    .split('')
    .map((character) => (character.charCodeAt(0) < 32 || '<>:"/\\|?*'.includes(character) ? '_' : character))
    .join('')
    .trim()
  if (!cleaned) return 'workbook.xlsx'
  return /\.xlsx$/i.test(cleaned) ? cleaned : `${cleaned.replace(/\.+$/, '')}.xlsx`
}

function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = normalizeFileNameInput(fileName)
  document.body.appendChild(link)
  link.click()
  link.remove()
  URL.revokeObjectURL(url)
}

function waitForNextFrame(): Promise<void> {
  return new Promise((resolve) => {
    requestAnimationFrame(() => resolve())
  })
}

function isEditableElement(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) {
    return false
  }

  if (target.isContentEditable) {
    return true
  }

  return !!target.closest('input, textarea, select, [contenteditable="true"], [role="textbox"]')
}

function isUniverEditorElement(target: EventTarget | null): boolean {
  if (!(target instanceof HTMLElement)) {
    return false
  }

  return !!target.closest('[data-editorid]')
}

function hasBrowserTextSelection(): boolean {
  const selection = window.getSelection()
  return !!selection && !selection.isCollapsed && selection.toString().trim().length > 0
}

function accelKey(event: KeyboardEvent): boolean {
  return event.ctrlKey || event.metaKey
}

function rangeContainsCell(
  range: { startRow: number; endRow: number; startColumn: number; endColumn: number },
  row: number,
  column: number,
): boolean {
  return row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn
}

function createPrimaryCell(range: {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}) {
  return {
    startRow: range.startRow,
    endRow: range.startRow,
    startColumn: range.startColumn,
    endColumn: range.startColumn,
    actualRow: range.startRow,
    actualColumn: range.startColumn,
    rangeType: 0,
    isMerged: false,
    isMergedMainCell: false,
  }
}

function getUniverService<T>(injector: { has(id: unknown): boolean; get(id: unknown): T }, id: unknown): T | null {
  if (!injector.has(id)) {
    return null
  }

  try {
    return injector.get(id)
  } catch {
    return null
  }
}

const workspaceFrameStyle: CSSProperties = {
  position: 'relative',
  width: '100%',
  height: '100%',
  overflow: 'hidden',
  borderRadius: 18,
  background: '#ffffff',
  border: '1px solid rgba(148, 163, 184, 0.22)',
  boxShadow: '0 18px 40px rgba(15, 23, 42, 0.12)',
  outline: 'none',
  display: 'flex',
  flexDirection: 'column',
}

const workspaceHostStyle: CSSProperties = {
  width: '100%',
  flex: 1,
  minHeight: 0,
  background: '#ffffff',
}

const fileToolbarStyle: CSSProperties = {
  flexShrink: 0,
  minHeight: 52,
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  maxWidth: '100%',
  borderBottom: '1px solid rgba(148, 163, 184, 0.18)',
  background: 'linear-gradient(180deg, #ffffff 0%, #f8fafc 100%)',
  padding: '8px 16px',
  pointerEvents: 'auto',
}

const fileGroupSelectStyle: CSSProperties = {
  height: 30,
  maxWidth: 180,
  border: '1px solid rgba(148, 163, 184, 0.26)',
  borderRadius: 8,
  background: '#ffffff',
  color: '#0f172a',
  padding: '0 9px',
  fontSize: 12,
}

const fileToolbarToggleStyle: CSSProperties = {
  height: 30,
  display: 'inline-flex',
  alignItems: 'center',
  gap: 6,
  border: '1px solid rgba(148, 163, 184, 0.24)',
  borderRadius: 999,
  background: '#ffffff',
  color: '#475569',
  padding: '0 10px',
  fontSize: 12,
  fontWeight: 700,
  cursor: 'pointer',
  whiteSpace: 'nowrap',
}

const userBadgeStyle: CSSProperties = {
  height: 30,
  display: 'inline-flex',
  alignItems: 'center',
  gap: 6,
  padding: '0 10px',
  border: '1px solid rgba(148, 163, 184, 0.24)',
  borderRadius: 999,
  background: '#ffffff',
  color: '#0f172a',
  fontSize: 12,
  fontWeight: 800,
  whiteSpace: 'nowrap',
}

const userRoleStyle: CSSProperties = {
  color: '#2563eb',
  fontSize: 11,
}

const fileToolbarCheckboxStyle: CSSProperties = {
  width: 13,
  height: 13,
  margin: 0,
  accentColor: '#2563eb',
}

const fileToolbarButtonStyle: CSSProperties = {
  height: 30,
  border: '1px solid rgba(37, 99, 235, 0.22)',
  borderRadius: 999,
  background: '#eff6ff',
  color: '#1d4ed8',
  padding: '0 12px',
  fontSize: 12,
  fontWeight: 800,
  cursor: 'pointer',
  whiteSpace: 'nowrap',
}

const fileToolbarDangerButtonStyle: CSSProperties = {
  ...fileToolbarButtonStyle,
  border: '1px solid rgba(220, 38, 38, 0.22)',
  background: '#fef2f2',
  color: '#b91c1c',
}

const fileStatusStyle: CSSProperties = {
  minWidth: 0,
  flex: 1,
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap',
  color: '#64748b',
  fontSize: 12,
}

const overlayShellStyle: CSSProperties = {
  position: 'absolute',
  top: 10,
  right: 16,
  zIndex: 30,
  pointerEvents: 'none',
}
