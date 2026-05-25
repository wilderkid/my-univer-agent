import { memo, useEffect, useState, type CSSProperties, type FormEvent } from 'react'
import { createPortal } from 'react-dom'
import { changePassword, type AuthUser } from '../lib/auth'

interface ChangePasswordButtonProps {
  buttonStyle: CSSProperties
  currentUser: AuthUser
}

export const ChangePasswordButton = memo(function ChangePasswordButton({ buttonStyle, currentUser }: ChangePasswordButtonProps) {
  const [open, setOpen] = useState(false)

  return (
    <>
      <button type="button" style={buttonStyle} onClick={() => setOpen(true)}>
        修改密码
      </button>
      {open ? <ChangePasswordDialog currentUser={currentUser} onClose={() => setOpen(false)} /> : null}
    </>
  )
})

interface ChangePasswordDialogProps {
  currentUser: AuthUser
  onClose: () => void
}

function ChangePasswordDialog({ currentUser, onClose }: ChangePasswordDialogProps) {
  const [password, setPassword] = useState('')
  const [confirmPassword, setConfirmPassword] = useState('')
  const [saving, setSaving] = useState(false)
  const [message, setMessage] = useState('')
  const [error, setError] = useState('')

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') onClose()
    }

    document.addEventListener('keydown', handleKeyDown)
    return () => document.removeEventListener('keydown', handleKeyDown)
  }, [onClose])

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()
    setMessage('')
    setError('')

    if (password.length < 6) {
      setError('新密码至少 6 位')
      return
    }

    if (password !== confirmPassword) {
      setError('两次输入的密码不一致')
      return
    }

    setSaving(true)
    try {
      await changePassword(password)
      setPassword('')
      setConfirmPassword('')
      setMessage('密码已修改，下次登录请使用新密码。')
    } catch (err) {
      setError(err instanceof Error ? err.message : '修改密码失败')
    } finally {
      setSaving(false)
    }
  }

  return createPortal(
    <div style={backdropStyle} onClick={onClose}>
      <form style={dialogStyle} onSubmit={(event) => void handleSubmit(event)} onClick={(event) => event.stopPropagation()}>
        <div style={headerStyle}>
          <div>
            <div style={eyebrowStyle}>{currentUser.username}</div>
            <h2 style={titleStyle}>修改密码</h2>
          </div>
          <button type="button" style={closeButtonStyle} onClick={onClose}>
            关闭
          </button>
        </div>

        <div style={bodyStyle}>
          <label style={fieldStyle}>
            <span style={labelStyle}>新密码</span>
            <input
              value={password}
              onChange={(event) => setPassword(event.target.value)}
              type="password"
              style={inputStyle}
              disabled={saving}
              autoFocus
            />
          </label>

          <label style={fieldStyle}>
            <span style={labelStyle}>确认新密码</span>
            <input
              value={confirmPassword}
              onChange={(event) => setConfirmPassword(event.target.value)}
              type="password"
              style={inputStyle}
              disabled={saving}
            />
          </label>

          {error ? <div style={errorStyle}>{error}</div> : null}
          {message ? <div style={successStyle}>{message}</div> : null}

          <button type="submit" style={primaryButtonStyle} disabled={saving || password.length < 6 || !confirmPassword}>
            {saving ? '正在保存...' : '保存新密码'}
          </button>
        </div>
      </form>
    </div>,
    document.body,
  )
}

const backdropStyle: CSSProperties = {
  position: 'fixed',
  inset: 0,
  zIndex: 65,
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  background: 'rgba(15, 23, 42, 0.52)',
  pointerEvents: 'auto',
}

const dialogStyle: CSSProperties = {
  width: 'min(420px, calc(100vw - 32px))',
  overflow: 'hidden',
  border: 0,
  borderRadius: 22,
  background: '#f8fafc',
  boxShadow: '0 28px 80px rgba(15, 23, 42, 0.28)',
}

const headerStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  padding: '20px 22px',
  background: 'linear-gradient(135deg, #0f172a 0%, #1e40af 100%)',
  color: '#ffffff',
}

const eyebrowStyle: CSSProperties = {
  fontSize: 11,
  fontWeight: 900,
  letterSpacing: '0.14em',
  color: '#bfdbfe',
}

const titleStyle: CSSProperties = {
  margin: '4px 0 0',
  fontSize: 22,
}

const closeButtonStyle: CSSProperties = {
  height: 32,
  border: '1px solid rgba(255,255,255,0.28)',
  borderRadius: 999,
  background: 'rgba(255,255,255,0.14)',
  color: '#ffffff',
  padding: '0 12px',
  fontWeight: 800,
  cursor: 'pointer',
}

const bodyStyle: CSSProperties = {
  display: 'grid',
  gap: 13,
  padding: 22,
}

const fieldStyle: CSSProperties = {
  display: 'grid',
  gap: 7,
}

const labelStyle: CSSProperties = {
  color: '#334155',
  fontSize: 12,
  fontWeight: 900,
}

const inputStyle: CSSProperties = {
  height: 38,
  border: '1px solid rgba(148, 163, 184, 0.38)',
  borderRadius: 11,
  padding: '0 11px',
  color: '#0f172a',
}

const primaryButtonStyle: CSSProperties = {
  height: 40,
  border: 0,
  borderRadius: 11,
  background: '#1d4ed8',
  color: '#ffffff',
  fontWeight: 900,
  cursor: 'pointer',
}

const errorStyle: CSSProperties = {
  padding: '9px 11px',
  borderRadius: 11,
  background: '#fef2f2',
  color: '#b91c1c',
  fontSize: 12,
}

const successStyle: CSSProperties = {
  ...errorStyle,
  background: '#dcfce7',
  color: '#166534',
}
