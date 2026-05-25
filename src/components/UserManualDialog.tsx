import { memo, useEffect, useState, type CSSProperties } from 'react'
import { createPortal } from 'react-dom'

interface UserManualDialogProps {
  onClose: () => void
}

interface UserManualButtonProps {
  buttonStyle: CSSProperties
}

export const UserManualButton = memo(function UserManualButton({ buttonStyle }: UserManualButtonProps) {
  const [open, setOpen] = useState(false)

  return (
    <>
      <button
        type="button"
        style={buttonStyle}
        onClick={() => setOpen(true)}
      >
        使用手册
      </button>
      {open ? <UserManualDialog onClose={() => setOpen(false)} /> : null}
    </>
  )
})

function UserManualDialog({ onClose }: UserManualDialogProps) {
  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        onClose()
      }
    }

    document.addEventListener('keydown', handleKeyDown)
    return () => document.removeEventListener('keydown', handleKeyDown)
  }, [onClose])

  return createPortal(
    <div style={manualBackdropStyle} onClick={onClose}>
      <div style={manualDialogStyle} onClick={(event) => event.stopPropagation()}>
        <div style={manualHeaderStyle}>
          <div>
            <div style={manualEyebrowStyle}>MY UNIVER AGENT</div>
            <h2 style={manualTitleStyle}>用户手册</h2>
          </div>
          <button type="button" style={manualCloseButtonStyle} onClick={onClose}>
            关闭
          </button>
        </div>

        <div style={manualBodyStyle}>
          <section style={manualSectionStyle}>
            <h3 style={manualSectionTitleStyle}>快速开始</h3>
            <ol style={manualListStyle}>
              <li>点击“导入 Excel”，可一次选择一个或多个 .xlsx 文件。</li>
              <li>文件导入后，顶部下拉框用于定位当前文件；切换 Univer 底部工作表时，下拉框会自动同步到该工作表所属文件。</li>
              <li>点击“导出 Excel”会导出当前下拉框选中的文件；点击“关闭文件”会关闭当前下拉框选中的文件。</li>
              <li>点击“清空”会清空当前已导入的文件列表和工作簿内容，使用前请先确认已经导出备份。</li>
            </ol>
          </section>

          <section style={manualSectionStyle}>
            <h3 style={manualSectionTitleStyle}>AI 面板使用</h3>
            <ol style={manualListStyle}>
              <li>先在 AI 配置中添加模型，再在 AI 面板中选择模型执行任务。</li>
              <li>输入 @ 可以选择工作表名，输入 # 可以绑定当前选区，格式类似“工作表名:A1:D10”。</li>
              <li>AI 可以读取和操作当前 Univer 中所有打开的工作表，包括跨文件、跨表处理。</li>
              <li>每次 AI 执行后，优先查看变更摘要和差异，再决定是否保留、撤销或再次应用。</li>
            </ol>
          </section>

          <section style={manualSectionStyle}>
            <h3 style={manualSectionTitleStyle}>推荐工作流</h3>
            <ol style={manualListStyle}>
              <li>导入文件后，先检查关键样式、公式、合并单元格是否正常。</li>
              <li>执行 AI 任务前，尽量用 @ 或 # 明确指定工作表或区域。</li>
              <li>对金额、数量、单价等关键数据，执行后必须人工抽查或让 AI 再做一次核对。</li>
              <li>完成阶段性处理后，及时导出 Excel 作为备份版本。</li>
            </ol>
          </section>

          <section style={manualWarningSectionStyle}>
            <h3 style={manualWarningTitleStyle}>使用禁忌</h3>
            <ul style={manualListStyle}>
              <li>不要把本系统当作唯一文件备份；重要文件必须保留原始 Excel，并定期导出结果。</li>
              <li>不要在未检查的情况下直接相信 AI 对金额、合同、发票、库存等关键数据的处理结果。</li>
              <li>不要导入过大的 Excel 文件；当前导入接口限制约 25MB，且超大表格会明显增加内存和等待时间。</li>
              <li>不要让 AI 随意合并单元格；合并可能隐藏非左上角单元格内容，只在明确需要版式调整时使用。</li>
              <li>不要在不可信网络中暴露本机后端服务；AI 配置中的 API Key 当前属于本机全局配置。</li>
              <li>当前版本还不是正式多人隔离系统，不建议多人共用同一个浏览器工作区处理不同私密文件。</li>
            </ul>
          </section>

          <section style={manualSectionStyle}>
            <h3 style={manualSectionTitleStyle}>当前限制</h3>
            <ul style={manualListStyle}>
              <li>Excel 导入/导出已覆盖基础数据、公式、样式、合并单元格，但复杂图表、宏、数据透视表不保证完整还原。</li>
              <li>AI 配置暂时是系统级配置，不是按用户隔离配置。</li>
              <li>多人账号、注册开关、用户文件隔离会在后续版本单独实现。</li>
            </ul>
          </section>
        </div>
      </div>
    </div>,
    document.body,
  )
}

const manualBackdropStyle: CSSProperties = {
  position: 'fixed',
  inset: 0,
  zIndex: 50,
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  background: 'rgba(15, 23, 42, 0.52)',
  pointerEvents: 'auto',
}

const manualDialogStyle: CSSProperties = {
  width: 'min(860px, calc(100vw - 32px))',
  maxHeight: 'min(820px, calc(100vh - 32px))',
  display: 'flex',
  flexDirection: 'column',
  overflow: 'hidden',
  borderRadius: 22,
  border: '1px solid rgba(148, 163, 184, 0.32)',
  background: '#f8fafc',
  boxShadow: '0 28px 80px rgba(15, 23, 42, 0.28)',
}

const manualHeaderStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  gap: 16,
  padding: '22px 24px',
  borderBottom: '1px solid rgba(148, 163, 184, 0.22)',
  background: 'linear-gradient(135deg, #0f172a 0%, #1e293b 58%, #334155 100%)',
  color: '#ffffff',
}

const manualEyebrowStyle: CSSProperties = {
  fontSize: 11,
  fontWeight: 900,
  letterSpacing: '0.16em',
  color: '#93c5fd',
}

const manualTitleStyle: CSSProperties = {
  margin: '4px 0 0',
  fontSize: 24,
  lineHeight: 1.2,
}

const manualCloseButtonStyle: CSSProperties = {
  height: 34,
  border: '1px solid rgba(255, 255, 255, 0.28)',
  borderRadius: 999,
  background: 'rgba(255, 255, 255, 0.12)',
  color: '#ffffff',
  padding: '0 14px',
  fontWeight: 800,
  cursor: 'pointer',
}

const manualBodyStyle: CSSProperties = {
  overflow: 'auto',
  padding: 24,
}

const manualSectionStyle: CSSProperties = {
  padding: '16px 18px',
  marginBottom: 14,
  border: '1px solid rgba(148, 163, 184, 0.24)',
  borderRadius: 16,
  background: '#ffffff',
}

const manualWarningSectionStyle: CSSProperties = {
  ...manualSectionStyle,
  borderColor: 'rgba(245, 158, 11, 0.38)',
  background: '#fffbeb',
}

const manualSectionTitleStyle: CSSProperties = {
  margin: '0 0 10px',
  color: '#0f172a',
  fontSize: 16,
}

const manualWarningTitleStyle: CSSProperties = {
  ...manualSectionTitleStyle,
  color: '#92400e',
}

const manualListStyle: CSSProperties = {
  margin: 0,
  paddingLeft: 20,
  color: '#334155',
  fontSize: 13,
  lineHeight: 1.8,
}
