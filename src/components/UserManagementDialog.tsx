import { memo, useEffect, useState, type CSSProperties, type FormEvent } from 'react'
import { createPortal } from 'react-dom'
import { createUser, listUsers, updateUser, type AdminUser, type AuthUser } from '../lib/auth'

interface UserManagementButtonProps {
  buttonStyle: CSSProperties
  currentUser: AuthUser
}

export const UserManagementButton = memo(function UserManagementButton({ buttonStyle, currentUser }: UserManagementButtonProps) {
  const [open, setOpen] = useState(false)

  if (currentUser.role !== 'admin') return null

  return (
    <>
      <button type="button" style={buttonStyle} onClick={() => setOpen(true)}>
        用户管理
      </button>
      {open ? <UserManagementDialog currentUser={currentUser} onClose={() => setOpen(false)} /> : null}
    </>
  )
})

interface UserManagementDialogProps {
  currentUser: AuthUser
  onClose: () => void
}

function UserManagementDialog({ currentUser, onClose }: UserManagementDialogProps) {
  const [users, setUsers] = useState<AdminUser[]>([])
  const [loading, setLoading] = useState(true)
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState('')
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const [role, setRole] = useState<AuthUser['role']>('user')

  async function refreshUsers() {
    setLoading(true)
    setError('')
    try {
      const result = await listUsers()
      setUsers(result.users)
    } catch (err) {
      setError(err instanceof Error ? err.message : '加载用户失败')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    let cancelled = false

    listUsers()
      .then((result) => {
        if (!cancelled) setUsers(result.users)
      })
      .catch((err: unknown) => {
        if (!cancelled) setError(err instanceof Error ? err.message : '加载用户失败')
      })
      .finally(() => {
        if (!cancelled) setLoading(false)
      })

    return () => {
      cancelled = true
    }
  }, [])

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') onClose()
    }

    document.addEventListener('keydown', handleKeyDown)
    return () => document.removeEventListener('keydown', handleKeyDown)
  }, [onClose])

  async function handleCreateUser(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()
    setSaving(true)
    setError('')

    try {
      await createUser({ username, password, role, isAdmin: role === 'admin' })
      setUsername('')
      setPassword('')
      setRole('user')
      await refreshUsers()
    } catch (err) {
      setError(err instanceof Error ? err.message : '创建用户失败')
    } finally {
      setSaving(false)
    }
  }

  async function handleToggleDisabled(user: AdminUser) {
    setSaving(true)
    setError('')
    try {
      await updateUser(user.id, { disabled: !user.disabled })
      await refreshUsers()
    } catch (err) {
      setError(err instanceof Error ? err.message : '更新用户状态失败')
    } finally {
      setSaving(false)
    }
  }

  async function handleChangeRole(user: AdminUser, nextRole: AuthUser['role']) {
    setSaving(true)
    setError('')
    try {
      await updateUser(user.id, { role: nextRole })
      await refreshUsers()
    } catch (err) {
      setError(err instanceof Error ? err.message : '更新角色失败')
    } finally {
      setSaving(false)
    }
  }

  async function handleResetPassword(user: AdminUser) {
    const nextPassword = window.prompt(`为 ${user.username} 设置新密码（至少 6 位）`)
    if (nextPassword === null) return

    setSaving(true)
    setError('')
    try {
      await updateUser(user.id, { password: nextPassword })
      await refreshUsers()
    } catch (err) {
      setError(err instanceof Error ? err.message : '重置密码失败')
    } finally {
      setSaving(false)
    }
  }

  return createPortal(
    <div style={backdropStyle} onClick={onClose}>
      <div style={dialogStyle} onClick={(event) => event.stopPropagation()}>
        <div style={headerStyle}>
          <div>
            <div style={eyebrowStyle}>ADMIN</div>
            <h2 style={titleStyle}>用户管理</h2>
          </div>
          <button type="button" style={closeButtonStyle} onClick={onClose}>
            关闭
          </button>
        </div>

        <div style={bodyStyle}>
          <form style={createFormStyle} onSubmit={(event) => void handleCreateUser(event)}>
            <div style={formTitleStyle}>创建账户</div>
            <input
              value={username}
              onChange={(event) => setUsername(event.target.value)}
              placeholder="用户名，如 zhangsan"
              style={inputStyle}
              disabled={saving}
            />
            <input
              value={password}
              onChange={(event) => setPassword(event.target.value)}
              placeholder="初始密码，至少 6 位"
              type="password"
              style={inputStyle}
              disabled={saving}
            />
            <div style={roleToggleStyle}>
              <button
                type="button"
                style={role === 'user' ? activeRoleButtonStyle : roleButtonStyle}
                onClick={() => setRole('user')}
                disabled={saving}
              >
                普通用户
              </button>
              <button
                type="button"
                style={role === 'admin' ? activeRoleButtonStyle : roleButtonStyle}
                onClick={() => setRole('admin')}
                disabled={saving}
              >
                管理员
              </button>
            </div>
            <button type="submit" style={primaryButtonStyle} disabled={saving || !username.trim() || password.length < 6}>
              创建
            </button>
          </form>

          {error ? <div style={errorStyle}>{error}</div> : null}

          <div style={tableShellStyle}>
            <div style={tableHeaderStyle}>
              <span>账户</span>
              <span>角色</span>
              <span>状态</span>
              <span>操作</span>
            </div>
            {loading ? (
              <div style={emptyStyle}>正在加载用户...</div>
            ) : users.length === 0 ? (
              <div style={emptyStyle}>暂无用户</div>
            ) : (
              users.map((user) => {
                const isSelf = user.id === currentUser.id
                return (
                  <div key={user.id} style={tableRowStyle}>
                    <div>
                      <div style={usernameStyle}>{user.username}</div>
                      <div style={metaStyle}>{new Date(user.createdAt).toLocaleString()}</div>
                    </div>
                    <button
                      type="button"
                      style={user.role === 'admin' ? adminRoleBadgeButtonStyle : userRoleBadgeButtonStyle}
                      onClick={() => void handleChangeRole(user, user.role === 'admin' ? 'user' : 'admin')}
                      disabled={saving || isSelf}
                      title={isSelf ? '不能修改自己的角色' : '点击切换角色'}
                    >
                      {user.role === 'admin' ? '管理员' : '普通用户'}
                    </button>
                    <span style={user.disabled ? disabledBadgeStyle : activeBadgeStyle}>
                      {user.disabled ? '已禁用' : '启用中'}
                    </span>
                    <div style={actionRowStyle}>
                      <button type="button" style={secondaryButtonStyle} onClick={() => void handleResetPassword(user)} disabled={saving}>
                        重置密码
                      </button>
                      <button
                        type="button"
                        style={user.disabled ? secondaryButtonStyle : dangerButtonStyle}
                        onClick={() => void handleToggleDisabled(user)}
                        disabled={saving || isSelf}
                      >
                        {user.disabled ? '启用' : '禁用'}
                      </button>
                    </div>
                  </div>
                )
              })
            )}
          </div>
        </div>
      </div>
    </div>,
    document.body,
  )
}

const backdropStyle: CSSProperties = {
  position: 'fixed',
  inset: 0,
  zIndex: 60,
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  background: 'rgba(15, 23, 42, 0.52)',
  pointerEvents: 'auto',
}

const dialogStyle: CSSProperties = {
  width: 'min(920px, calc(100vw - 32px))',
  maxHeight: 'min(820px, calc(100vh - 32px))',
  display: 'flex',
  flexDirection: 'column',
  overflow: 'hidden',
  borderRadius: 22,
  background: '#f8fafc',
  boxShadow: '0 28px 80px rgba(15, 23, 42, 0.28)',
}

const headerStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  padding: '22px 24px',
  background: 'linear-gradient(135deg, #172554 0%, #1d4ed8 100%)',
  color: '#ffffff',
}

const eyebrowStyle: CSSProperties = {
  fontSize: 11,
  fontWeight: 900,
  letterSpacing: '0.16em',
  color: '#bfdbfe',
}

const titleStyle: CSSProperties = {
  margin: '4px 0 0',
  fontSize: 24,
}

const closeButtonStyle: CSSProperties = {
  height: 34,
  border: '1px solid rgba(255,255,255,0.28)',
  borderRadius: 999,
  background: 'rgba(255,255,255,0.14)',
  color: '#ffffff',
  padding: '0 14px',
  fontWeight: 800,
  cursor: 'pointer',
}

const bodyStyle: CSSProperties = {
  overflow: 'auto',
  padding: 22,
}

const createFormStyle: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: '1fr 1fr 150px auto',
  gap: 10,
  alignItems: 'end',
  padding: 14,
  border: '1px solid rgba(148, 163, 184, 0.24)',
  borderRadius: 16,
  background: '#ffffff',
}

const formTitleStyle: CSSProperties = {
  gridColumn: '1 / -1',
  color: '#0f172a',
  fontWeight: 900,
}

const inputStyle: CSSProperties = {
  height: 36,
  border: '1px solid rgba(148, 163, 184, 0.36)',
  borderRadius: 10,
  padding: '0 10px',
}

const roleToggleStyle: CSSProperties = {
  height: 36,
  display: 'grid',
  gridTemplateColumns: '1fr 1fr',
  gap: 4,
  padding: 3,
  border: '1px solid rgba(148, 163, 184, 0.32)',
  borderRadius: 10,
  background: '#f8fafc',
}

const roleButtonStyle: CSSProperties = {
  border: 0,
  borderRadius: 8,
  background: 'transparent',
  color: '#475569',
  fontWeight: 900,
  cursor: 'pointer',
}

const activeRoleButtonStyle: CSSProperties = {
  ...roleButtonStyle,
  background: '#1d4ed8',
  color: '#ffffff',
}

const userRoleBadgeButtonStyle: CSSProperties = {
  width: 96,
  height: 30,
  border: '1px solid rgba(148, 163, 184, 0.34)',
  borderRadius: 999,
  background: '#f8fafc',
  color: '#475569',
  fontWeight: 900,
  cursor: 'pointer',
}

const adminRoleBadgeButtonStyle: CSSProperties = {
  ...userRoleBadgeButtonStyle,
  borderColor: 'rgba(37, 99, 235, 0.32)',
  background: '#dbeafe',
  color: '#1d4ed8',
}

const primaryButtonStyle: CSSProperties = {
  height: 36,
  border: 0,
  borderRadius: 10,
  background: '#1d4ed8',
  color: '#ffffff',
  padding: '0 14px',
  fontWeight: 900,
  cursor: 'pointer',
}

const secondaryButtonStyle: CSSProperties = {
  height: 30,
  border: '1px solid rgba(148, 163, 184, 0.34)',
  borderRadius: 9,
  background: '#ffffff',
  color: '#0f172a',
  padding: '0 10px',
  fontWeight: 800,
  cursor: 'pointer',
}

const dangerButtonStyle: CSSProperties = {
  ...secondaryButtonStyle,
  borderColor: 'rgba(239, 68, 68, 0.36)',
  color: '#b91c1c',
}

const errorStyle: CSSProperties = {
  marginTop: 12,
  padding: '10px 12px',
  borderRadius: 12,
  background: '#fef2f2',
  color: '#b91c1c',
  fontSize: 12,
}

const tableShellStyle: CSSProperties = {
  marginTop: 16,
  overflow: 'hidden',
  border: '1px solid rgba(148, 163, 184, 0.24)',
  borderRadius: 16,
  background: '#ffffff',
}

const tableHeaderStyle: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: '1.5fr 130px 100px 220px',
  gap: 12,
  padding: '11px 14px',
  background: '#f1f5f9',
  color: '#475569',
  fontSize: 12,
  fontWeight: 900,
}

const tableRowStyle: CSSProperties = {
  display: 'grid',
  gridTemplateColumns: '1.5fr 130px 100px 220px',
  gap: 12,
  alignItems: 'center',
  padding: '12px 14px',
  borderTop: '1px solid rgba(148, 163, 184, 0.18)',
}

const usernameStyle: CSSProperties = {
  color: '#0f172a',
  fontWeight: 900,
}

const metaStyle: CSSProperties = {
  marginTop: 3,
  color: '#64748b',
  fontSize: 11,
}

const activeBadgeStyle: CSSProperties = {
  width: 'fit-content',
  borderRadius: 999,
  background: '#dcfce7',
  color: '#166534',
  padding: '4px 8px',
  fontSize: 11,
  fontWeight: 900,
}

const disabledBadgeStyle: CSSProperties = {
  ...activeBadgeStyle,
  background: '#fee2e2',
  color: '#991b1b',
}

const actionRowStyle: CSSProperties = {
  display: 'flex',
  gap: 8,
}

const emptyStyle: CSSProperties = {
  padding: 18,
  color: '#64748b',
  fontSize: 13,
}
