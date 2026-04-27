import { useState, useRef, useEffect, type KeyboardEvent } from 'react'
import type { UniverAgent } from '../lib/univer-agent'
import type { LLMProvider } from '../lib/agent'
import { initAgent, getAgent, destroyAgent } from '../lib/agent'
import { saveAIConfig, loadAIConfig } from '../lib/storage'

interface LogEntry {
  id: number
  type: 'cmd' | 'ok' | 'err' | 'running'
  text: string
}

const PRESET_CMDS = [
  '在A1输入"姓名"，B1输入"成绩"，A2输入"张三"，B2输入90',
  '将第1行（A1:B1）字体加粗并设置背景色为浅蓝色',
  '在C2输入公式 =B2*1.1',
  '新建一个工作表，命名为"数据汇总"',
  '把A1:C3区域合并单元格',
]

const PROVIDER_LABELS: Record<LLMProvider, string> = {
  qwen: '通义千问',
  openai: 'OpenAI',
  custom: '自定义',
}

const PROVIDER_MODEL_PLACEHOLDER: Record<LLMProvider, string> = {
  qwen: '默认 qwen-plus',
  openai: '默认 gpt-4o-mini',
  custom: '模型名称（必填）',
}

const VALID_PROVIDERS: LLMProvider[] = ['qwen', 'openai', 'custom']

export function AgentPanel() {
  // Resolve initial config: env vars > localStorage > defaults
  const [apiKey, setApiKey] = useState(() => {
    try {
      const envKey = import.meta.env.VITE_LLM_API_KEY as string | undefined
      return envKey || loadAIConfig()?.apiKey || ''
    } catch { return '' }
  })
  const [provider, setProvider] = useState<LLMProvider>(() => {
    try {
      const envP = import.meta.env.VITE_LLM_PROVIDER as string | undefined
      const savedP = loadAIConfig()?.provider
      const resolved = envP || savedP || 'qwen'
      return VALID_PROVIDERS.includes(resolved as LLMProvider)
        ? (resolved as LLMProvider)
        : 'qwen'
    } catch { return 'qwen' }
  })
  const [model, setModel] = useState(() => {
    try {
      const envM = import.meta.env.VITE_LLM_MODEL as string | undefined
      return envM || loadAIConfig()?.model || ''
    } catch { return '' }
  })
  const [baseURL, setBaseURL] = useState(() => {
    try {
      const envB = import.meta.env.VITE_LLM_BASE_URL as string | undefined
      return envB || loadAIConfig()?.baseURL || ''
    } catch { return '' }
  })
  const [configured, setConfigured] = useState(false)
  const [cmd, setCmd] = useState('')
  const [running, setRunning] = useState(false)
  const [logs, setLogs] = useState<LogEntry[]>([])
  const [collapsed, setCollapsed] = useState(false)
  const logIdRef = useRef(0)
  const logEndRef = useRef<HTMLDivElement>(null)
  const agentRef = useRef<UniverAgent | null>(null)

  useEffect(() => {
    if (apiKey && !configured) {
      handleConfigure()
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  useEffect(() => {
    logEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [logs])

  function addLog(type: LogEntry['type'], text: string) {
    setLogs(prev => [...prev, { id: logIdRef.current++, type, text }])
  }

  function handleConfigure() {
    if (!apiKey.trim()) return
    if (provider === 'custom' && (!baseURL.trim() || !model.trim())) return

    agentRef.current = initAgent({
      apiKey: apiKey.trim(),
      provider,
      model: model.trim() || undefined,
      baseURL: baseURL.trim() || undefined,
    })

    // Persist config so it survives page refresh
    saveAIConfig({
      provider,
      apiKey: apiKey.trim(),
      model: model.trim(),
      baseURL: baseURL.trim(),
    })

    setConfigured(true)
    const label = provider === 'custom'
      ? `自定义 (${model.trim()})`
      : PROVIDER_LABELS[provider]
    addLog('ok', `已连接 ${label}`)
  }

  async function handleExecute() {
    const agent = agentRef.current ?? getAgent()
    if (!agent || !cmd.trim() || running) return

    const instruction = cmd.trim()
    setCmd('')
    setRunning(true)
    addLog('cmd', instruction)
    addLog('running', '执行中…')

    try {
      const result = await agent.execute(instruction)
      setLogs(prev => prev.filter(l => l.type !== 'running'))
      addLog('ok', result || '完成')
    } catch (e) {
      setLogs(prev => prev.filter(l => l.type !== 'running'))
      addLog('err', e instanceof Error ? e.message : '执行失败')
    } finally {
      setRunning(false)
    }
  }

  function handleKeyDown(e: KeyboardEvent<HTMLInputElement>) {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault()
      handleExecute()
    }
  }

  const canConfigure =
    !!apiKey.trim() &&
    (provider !== 'custom' || (!!baseURL.trim() && !!model.trim()))

  const logColor: Record<LogEntry['type'], string> = {
    cmd: '#60a5fa',
    ok: '#4ade80',
    err: '#f87171',
    running: '#facc15',
  }

  return (
    <div style={panelStyle}>
      {/* Header */}
      <div
        style={headerStyle}
        onClick={() => setCollapsed(c => !c)}
        title="点击折叠/展开"
      >
        <span style={{ fontWeight: 700, fontSize: 13 }}>🤖 AI 助手</span>
        <span style={{ fontSize: 11, opacity: 0.7 }}>
          {running ? '⏳ 执行中' : configured ? '● 已连接' : '○ 未配置'}
        </span>
        <span style={{ marginLeft: 'auto', fontSize: 12 }}>{collapsed ? '▲' : '▼'}</span>
      </div>

      {!collapsed && (
        <div style={bodyStyle}>
          {!configured ? (
            <div style={sectionStyle}>
              {/* Provider selector */}
              <div style={labelStyle}>LLM 提供商</div>
              <div style={{ display: 'flex', gap: 4, marginBottom: 10, flexWrap: 'wrap' }}>
                {(Object.keys(PROVIDER_LABELS) as LLMProvider[]).map(p => (
                  <button
                    key={p}
                    style={{
                      ...tagBtn,
                      background: provider === p ? '#3b82f6' : '#374151',
                      color: provider === p ? '#fff' : '#d1d5db',
                    }}
                    onClick={() => setProvider(p)}
                  >
                    {PROVIDER_LABELS[p]}
                  </button>
                ))}
              </div>

              {/* Custom Base URL — only for custom provider */}
              {provider === 'custom' && (
                <>
                  <div style={labelStyle}>API 地址</div>
                  <input
                    type="url"
                    placeholder="http://localhost:11434/v1"
                    value={baseURL}
                    onChange={e => setBaseURL(e.target.value)}
                    style={{ ...inputStyle, marginBottom: 8 }}
                  />
                </>
              )}

              {/* Model — optional for presets, required for custom */}
              <div style={labelStyle}>
                模型{provider !== 'custom' && <span style={{ opacity: 0.5 }}>（可选）</span>}
              </div>
              <input
                type="text"
                placeholder={PROVIDER_MODEL_PLACEHOLDER[provider]}
                value={model}
                onChange={e => setModel(e.target.value)}
                style={{ ...inputStyle, marginBottom: 8 }}
              />

              {/* API Key */}
              <div style={labelStyle}>API Key</div>
              <input
                type="password"
                placeholder="输入你的 API Key…"
                value={apiKey}
                onChange={e => setApiKey(e.target.value)}
                style={{ ...inputStyle, marginBottom: 8 }}
              />

              <button
                style={{
                  ...btnStyle,
                  width: '100%',
                  opacity: canConfigure ? 1 : 0.4,
                  cursor: canConfigure ? 'pointer' : 'not-allowed',
                }}
                onClick={handleConfigure}
                disabled={!canConfigure}
              >
                连接
              </button>
            </div>
          ) : (
            <>
              {/* Log area */}
              <div style={logAreaStyle}>
                {logs.length === 0 && (
                  <div style={{ color: '#6b7280', fontSize: 11, textAlign: 'center', paddingTop: 12 }}>
                    输入指令，AI 将操作表格
                  </div>
                )}
                {logs.map(l => (
                  <div key={l.id} style={{ color: logColor[l.type], fontSize: 12, marginBottom: 3, wordBreak: 'break-all' }}>
                    {l.type === 'cmd' ? '› ' : l.type === 'ok' ? '✓ ' : l.type === 'err' ? '✗ ' : '⟳ '}
                    {l.text}
                  </div>
                ))}
                <div ref={logEndRef} />
              </div>

              {/* Preset commands */}
              <div style={{ marginBottom: 8 }}>
                <div style={labelStyle}>快捷指令</div>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                  {PRESET_CMDS.map(c => (
                    <button
                      key={c}
                      style={tagBtn}
                      onClick={() => setCmd(c)}
                    >
                      {c}
                    </button>
                  ))}
                </div>
              </div>

              {/* Input */}
              <div style={{ display: 'flex', gap: 6 }}>
                <input
                  value={cmd}
                  onChange={e => setCmd(e.target.value)}
                  onKeyDown={handleKeyDown}
                  placeholder="输入自然语言指令…"
                  disabled={running}
                  style={{ ...inputStyle, flex: 1 }}
                />
                <button
                  style={{ ...btnStyle, padding: '0 12px' }}
                  onClick={handleExecute}
                  disabled={running || !cmd.trim()}
                >
                  {running ? '…' : '执行'}
                </button>
              </div>

              <button
                style={{ ...tagBtn, marginTop: 8, fontSize: 10, opacity: 0.5 }}
                onClick={() => { setConfigured(false); agentRef.current = null; destroyAgent() }}
              >
                重新配置
              </button>
            </>
          )}
        </div>
      )}
    </div>
  )
}

// ─── Inline styles ────────────────────────────────────────────────────────────

const panelStyle: React.CSSProperties = {
  position: 'fixed',
  bottom: 20,
  right: 20,
  width: 300,
  zIndex: 9999,
  background: '#1f2937',
  border: '1px solid #374151',
  borderRadius: 12,
  boxShadow: '0 8px 32px rgba(0,0,0,0.5)',
  color: '#f3f4f6',
  fontFamily: 'system-ui, sans-serif',
  overflow: 'hidden',
}

const headerStyle: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  padding: '10px 14px',
  background: '#111827',
  cursor: 'pointer',
  userSelect: 'none',
}

const bodyStyle: React.CSSProperties = {
  padding: '12px 14px',
}

const sectionStyle: React.CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
}

const labelStyle: React.CSSProperties = {
  fontSize: 11,
  color: '#9ca3af',
  marginBottom: 4,
  textTransform: 'uppercase',
  letterSpacing: '0.05em',
}

const inputStyle: React.CSSProperties = {
  background: '#111827',
  border: '1px solid #374151',
  borderRadius: 6,
  color: '#f3f4f6',
  padding: '6px 10px',
  fontSize: 12,
  outline: 'none',
  width: '100%',
  boxSizing: 'border-box',
}

const btnStyle: React.CSSProperties = {
  background: '#3b82f6',
  color: '#fff',
  border: 'none',
  borderRadius: 6,
  padding: '7px 14px',
  fontSize: 12,
  cursor: 'pointer',
  fontWeight: 600,
}

const tagBtn: React.CSSProperties = {
  background: '#374151',
  color: '#d1d5db',
  border: 'none',
  borderRadius: 4,
  padding: '3px 8px',
  fontSize: 11,
  cursor: 'pointer',
}

const logAreaStyle: React.CSSProperties = {
  background: '#111827',
  borderRadius: 6,
  padding: '8px 10px',
  minHeight: 80,
  maxHeight: 160,
  overflowY: 'auto',
  marginBottom: 10,
  fontFamily: 'monospace',
}
