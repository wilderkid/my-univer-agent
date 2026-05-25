export interface BackendAIModel {
  id: string
  provider: string
  model: string
  label: string
  apiKey: string
  baseURL: string
  createdAt: string
  updatedAt: string
}

export interface BackendAISettings {
  activeModelId: string
  models: BackendAIModel[]
}

export interface DiscoverAIModelsResult {
  baseURL: string
  models: string[]
}

export interface DiscoverAIModelsPayload {
  provider: string
  apiKey: string
  baseURL: string
}

export interface SaveAIModelsPayload {
  provider: string
  apiKey: string
  baseURL: string
  models: string[]
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
      // Ignore non-JSON error responses.
    }
    throw new Error(message)
  }

  if (response.status === 204) {
    return undefined as T
  }

  return response.json() as Promise<T>
}

export function loadBackendAISettings() {
  return request<BackendAISettings>('/api/settings/ai')
}

export function discoverBackendAIModels(payload: DiscoverAIModelsPayload) {
  return request<DiscoverAIModelsResult>('/api/settings/ai/discover-models', {
    method: 'POST',
    body: JSON.stringify(payload),
  })
}

export function saveBackendAIModels(payload: SaveAIModelsPayload) {
  return request<BackendAISettings>('/api/settings/ai/models', {
    method: 'POST',
    body: JSON.stringify(payload),
  })
}

export function selectActiveBackendAIModel(modelId: string) {
  return request<BackendAISettings>('/api/settings/ai/active-model', {
    method: 'PUT',
    body: JSON.stringify({ modelId }),
  })
}

export function deleteBackendAIModel(modelId: string) {
  return request<BackendAISettings>(`/api/settings/ai/models/${encodeURIComponent(modelId)}`, {
    method: 'DELETE',
  })
}

export function clearBackendAIConfig() {
  return request<void>('/api/settings/ai', {
    method: 'DELETE',
  })
}
