export interface AuthUser {
  id: string
  username: string
  role: 'admin' | 'user'
}

export interface AdminUser extends AuthUser {
  disabled: boolean
  createdAt: string
  updatedAt: string
}

export interface AuthState {
  authenticated: boolean
  user: AuthUser | null
}

const API_BASE = import.meta.env.VITE_APP_API_BASE ?? ''

async function request<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(`${API_BASE}${path}`, {
    credentials: 'include',
    headers: {
      'Content-Type': 'application/json',
      ...(init?.headers ?? {}),
    },
    ...init,
  })

  if (!response.ok) {
    let message = `Request failed with status ${response.status}`
    try {
      const data = await response.json() as { error?: string }
      if (data.error) message = data.error
    } catch {
      // Ignore non-JSON responses.
    }
    throw new Error(message)
  }

  if (response.status === 204) {
    return undefined as T
  }

  return response.json() as Promise<T>
}

export function getCurrentUser() {
  return request<AuthState>('/api/auth/me')
}

export function login(username: string, password: string) {
  return request<{ user: AuthUser }>('/api/auth/login', {
    method: 'POST',
    body: JSON.stringify({ username, password }),
  })
}

export function logout() {
  return request<void>('/api/auth/logout', {
    method: 'POST',
  })
}

export function changePassword(password: string) {
  return request<{ user: AuthUser }>('/api/auth/password', {
    method: 'PUT',
    body: JSON.stringify({ password }),
  })
}

export function listUsers() {
  return request<{ users: AdminUser[] }>('/api/admin/users')
}

export function createUser(payload: { username: string; password: string; role: AuthUser['role']; isAdmin?: boolean }) {
  return request<{ user: AdminUser }>('/api/admin/users', {
    method: 'POST',
    body: JSON.stringify(payload),
  })
}

export function updateUser(userId: string, payload: Partial<{ password: string; role: AuthUser['role']; disabled: boolean }>) {
  return request<{ user: AdminUser }>(`/api/admin/users/${encodeURIComponent(userId)}`, {
    method: 'PUT',
    body: JSON.stringify(payload),
  })
}
