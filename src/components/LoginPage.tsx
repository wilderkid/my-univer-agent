import { useState, type CSSProperties, type FormEvent } from 'react'

interface LoginPageProps {
  loading?: boolean
  error?: string
  onLogin: (username: string, password: string) => Promise<void>
}

export function LoginPage({ loading = false, error = '', onLogin }: LoginPageProps) {
  const [username, setUsername] = useState('admin')
  const [password, setPassword] = useState('')

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()
    await onLogin(username, password)
  }

  return (
    <div style={loginShellStyle}>
      <form style={loginCardStyle} onSubmit={(event) => void handleSubmit(event)}>
        <div style={loginEyebrowStyle}>MY UNIVER AGENT</div>
        <h1 style={loginTitleStyle}>登录系统</h1>
        <p style={loginDescriptionStyle}>当前系统不开放注册，仅管理员创建的账户可以登录。</p>

        <label style={fieldStyle}>
          <span style={labelStyle}>用户名</span>
          <input
            value={username}
            onChange={(event) => setUsername(event.target.value)}
            style={inputStyle}
            autoComplete="username"
            disabled={loading}
          />
        </label>

        <label style={fieldStyle}>
          <span style={labelStyle}>密码</span>
          <input
            value={password}
            onChange={(event) => setPassword(event.target.value)}
            style={inputStyle}
            type="password"
            autoComplete="current-password"
            disabled={loading}
          />
        </label>

        {error ? <div style={errorStyle}>{error}</div> : null}

        <button type="submit" style={submitButtonStyle} disabled={loading || !username.trim() || !password}>
          {loading ? '正在登录...' : '登录'}
        </button>
      </form>
    </div>
  )
}

const loginShellStyle: CSSProperties = {
  position: 'relative',
  zIndex: 1,
  width: '100%',
  height: '100%',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  padding: 24,
}

const loginCardStyle: CSSProperties = {
  width: 'min(420px, 100%)',
  padding: 28,
  borderRadius: 24,
  border: '1px solid rgba(148, 163, 184, 0.28)',
  background: 'rgba(255, 255, 255, 0.92)',
  boxShadow: '0 24px 70px rgba(15, 23, 42, 0.2)',
}

const loginEyebrowStyle: CSSProperties = {
  fontSize: 11,
  fontWeight: 900,
  letterSpacing: '0.16em',
  color: '#2563eb',
}

const loginTitleStyle: CSSProperties = {
  margin: '8px 0 8px',
  color: '#0f172a',
  fontSize: 28,
}

const loginDescriptionStyle: CSSProperties = {
  margin: '0 0 22px',
  color: '#475569',
  fontSize: 13,
  lineHeight: 1.7,
}

const fieldStyle: CSSProperties = {
  display: 'grid',
  gap: 7,
  marginBottom: 14,
}

const labelStyle: CSSProperties = {
  color: '#334155',
  fontSize: 12,
  fontWeight: 800,
}

const inputStyle: CSSProperties = {
  height: 40,
  border: '1px solid rgba(148, 163, 184, 0.42)',
  borderRadius: 12,
  padding: '0 12px',
  color: '#0f172a',
  background: '#ffffff',
  outline: 'none',
}

const errorStyle: CSSProperties = {
  margin: '4px 0 14px',
  padding: '10px 12px',
  borderRadius: 12,
  background: '#fef2f2',
  color: '#b91c1c',
  fontSize: 12,
}

const submitButtonStyle: CSSProperties = {
  width: '100%',
  height: 42,
  border: 0,
  borderRadius: 12,
  background: '#1d4ed8',
  color: '#ffffff',
  fontWeight: 900,
  cursor: 'pointer',
}
