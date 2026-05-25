import { useEffect, useState, type CSSProperties } from 'react'
import { AgentPanel } from './components/AgentPanel'
import { LoginPage } from './components/LoginPage'
import { UniverWorkspace } from './components/UniverWorkspace'
import { getCurrentUser, login, logout, type AuthUser } from './lib/auth'

export default function App() {
  const [user, setUser] = useState<AuthUser | null>(null)
  const [checkingAuth, setCheckingAuth] = useState(true)
  const [loggingIn, setLoggingIn] = useState(false)
  const [authError, setAuthError] = useState('')

  useEffect(() => {
    let cancelled = false

    async function checkAuth() {
      try {
        const state = await getCurrentUser()
        if (!cancelled) setUser(state.user)
      } catch {
        if (!cancelled) setUser(null)
      } finally {
        if (!cancelled) setCheckingAuth(false)
      }
    }

    void checkAuth()
    return () => {
      cancelled = true
    }
  }, [])

  async function handleLogin(username: string, password: string) {
    setLoggingIn(true)
    setAuthError('')
    try {
      const result = await login(username, password)
      setUser(result.user)
    } catch (error) {
      setAuthError(error instanceof Error ? error.message : '登录失败')
    } finally {
      setLoggingIn(false)
    }
  }

  async function handleLogout() {
    await logout()
    setUser(null)
  }

  return (
    <div style={appShellStyle}>
      <div style={backdropStyle} />
      {checkingAuth ? (
        <div style={loadingStyle}>正在检查登录状态...</div>
      ) : user ? (
        <div style={workspaceShellStyle}>
          <UniverWorkspace
            currentUser={user}
            onLogout={() => void handleLogout()}
            overlay={<AgentPanel currentUser={user} />}
          />
        </div>
      ) : (
        <LoginPage loading={loggingIn} error={authError} onLogin={handleLogin} />
      )}
    </div>
  )
}

const appShellStyle: CSSProperties = {
  position: 'relative',
  width: '100%',
  height: '100%',
  overflow: 'hidden',
  background: '#dbe4f0',
}

const backdropStyle: CSSProperties = {
  position: 'absolute',
  inset: 0,
  background: 'radial-gradient(circle at top left, #e2e8f0 0%, #dbeafe 45%, #cbd5e1 100%)',
}

const workspaceShellStyle: CSSProperties = {
  position: 'relative',
  zIndex: 1,
  width: '100%',
  height: '100%',
  padding: 12,
}

const loadingStyle: CSSProperties = {
  position: 'relative',
  zIndex: 1,
  width: '100%',
  height: '100%',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  color: '#0f172a',
  fontWeight: 800,
}
