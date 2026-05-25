import { randomBytes, randomUUID, scryptSync, timingSafeEqual } from 'node:crypto'
import { execFile } from 'node:child_process'
import { createReadStream } from 'node:fs'
import { promises as fs } from 'node:fs'
import { createServer } from 'node:http'
import { promisify } from 'node:util'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import ExcelJS from 'exceljs'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)
const projectRoot = path.resolve(__dirname, '..')
const distRoot = path.join(projectRoot, 'dist')
const dataRoot = path.join(projectRoot, '.data')
const usersRoot = path.join(dataRoot, 'users')
const usersPath = path.join(dataRoot, 'users.json')
const sessionsPath = path.join(dataRoot, 'sessions.json')
const legacyAIConfigPath = path.join(dataRoot, 'ai-config.json')
const adminUserId = 'admin'
const aiConfigPath = getUserAIConfigPath(adminUserId)

const host = process.env.APP_SERVER_HOST ?? '127.0.0.1'
const port = Number(process.env.APP_SERVER_PORT ?? 8787)
const execFileAsync = promisify(execFile)

const defaultAISettings = {
  activeModelId: '',
  models: [],
}

const defaultAdminUser = {
  id: adminUserId,
  username: 'admin',
  role: 'admin',
  disabled: false,
}

const sessionMaxAgeMs = 7 * 24 * 60 * 60 * 1000

const maxImportCells = 200000
const maxColorDebugLogs = Number(process.env.APP_XLSX_COLOR_DEBUG_LIMIT ?? 80)

const providerDefaults = {
  qwen: {
    baseURL: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
  },
  openai: {
    baseURL: 'https://api.openai.com/v1',
  },
}

const mimeTypes = {
  '.css': 'text/css; charset=utf-8',
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.txt': 'text/plain; charset=utf-8',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
}

function setCorsHeaders(request, response) {
  const origin = request.headers.origin
  response.setHeader('Access-Control-Allow-Origin', origin || '*')
  if (origin) response.setHeader('Vary', 'Origin')
  response.setHeader('Access-Control-Allow-Methods', 'GET,POST,PUT,DELETE,OPTIONS')
  response.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-File-Name, X-Import-Diagnostics')
  response.setHeader('Access-Control-Allow-Credentials', 'true')
}

function sendJson(response, statusCode, payload) {
  response.writeHead(statusCode, { 'Content-Type': 'application/json; charset=utf-8' })
  response.end(JSON.stringify(payload))
}

function sendError(response, statusCode, message) {
  sendJson(response, statusCode, { error: message })
}

function getUserAIConfigPath(userId) {
  return path.join(usersRoot, userId, 'ai-config.json')
}

function publicUser(user) {
  if (!user) return null
  return {
    id: user.id,
    username: user.username,
    role: user.role,
  }
}

function publicAdminUser(user) {
  if (!user) return null
  return {
    ...publicUser(user),
    disabled: user.disabled === true,
    createdAt: user.createdAt,
    updatedAt: user.updatedAt,
  }
}

function normalizeUsername(value) {
  return typeof value === 'string' ? value.trim() : ''
}

function isValidUsername(username) {
  return /^[A-Za-z0-9_.-]{2,32}$/.test(username)
}

function normalizeUserRole(value) {
  return value === 'admin' ? 'admin' : 'user'
}

function hashPassword(password, salt = randomBytes(16).toString('hex')) {
  const hash = scryptSync(String(password), salt, 64).toString('hex')
  return `scrypt:${salt}:${hash}`
}

function verifyPassword(password, passwordHash) {
  const parts = String(passwordHash || '').split(':')
  if (parts.length !== 3 || parts[0] !== 'scrypt') return false

  const expected = Buffer.from(parts[2], 'hex')
  const actual = scryptSync(String(password), parts[1], expected.length)
  return expected.length === actual.length && timingSafeEqual(expected, actual)
}

async function readJsonFile(filePath, fallback) {
  try {
    return JSON.parse(await fs.readFile(filePath, 'utf8'))
  } catch {
    return fallback
  }
}

async function writeJsonFileAtomic(filePath, value) {
  await fs.mkdir(path.dirname(filePath), { recursive: true })
  const tempPath = `${filePath}.${process.pid}.${Date.now()}.tmp`
  await fs.writeFile(tempPath, JSON.stringify(value, null, 2), 'utf8')
  await fs.rename(tempPath, filePath)
}

async function loadUsers() {
  const parsed = await readJsonFile(usersPath, null)
  const users = Array.isArray(parsed?.users) ? parsed.users : []
  const hasAdmin = users.some((user) => user?.username === defaultAdminUser.username)

  if (hasAdmin) {
    return users.map(normalizeUser).filter(Boolean)
  }

  const nextUsers = [
    ...users.map(normalizeUser).filter(Boolean),
    {
      ...defaultAdminUser,
      passwordHash: hashPassword('123456'),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    },
  ]
  await saveUsers(nextUsers)
  return nextUsers
}

function normalizeUser(user) {
  if (!user || typeof user !== 'object') return null
  const username = typeof user.username === 'string' ? user.username.trim() : ''
  const id = typeof user.id === 'string' && user.id.trim() ? user.id.trim() : username
  if (!username || !id || typeof user.passwordHash !== 'string') return null

  return {
    id,
    username,
    passwordHash: user.passwordHash,
    role: user.role === 'admin' ? 'admin' : 'user',
    disabled: user.disabled === true,
    createdAt: typeof user.createdAt === 'string' ? user.createdAt : new Date().toISOString(),
    updatedAt: typeof user.updatedAt === 'string' ? user.updatedAt : new Date().toISOString(),
  }
}

async function saveUsers(users) {
  await writeJsonFileAtomic(usersPath, { users })
}

async function loadSessions() {
  const parsed = await readJsonFile(sessionsPath, null)
  const now = Date.now()
  const sessions = Array.isArray(parsed?.sessions) ? parsed.sessions : []
  return sessions.filter((session) =>
    session &&
    typeof session.id === 'string' &&
    typeof session.userId === 'string' &&
    typeof session.expiresAt === 'string' &&
    Date.parse(session.expiresAt) > now
  )
}

async function saveSessions(sessions) {
  await writeJsonFileAtomic(sessionsPath, { sessions })
}

function parseCookies(request) {
  const header = request.headers.cookie
  const cookies = new Map()
  if (!header) return cookies

  for (const part of header.split(';')) {
    const index = part.indexOf('=')
    if (index === -1) continue
    const key = part.slice(0, index).trim()
    const value = part.slice(index + 1).trim()
    cookies.set(key, decodeURIComponent(value))
  }

  return cookies
}

function setSessionCookie(response, sessionId) {
  response.setHeader('Set-Cookie', [
    `univer_session=${encodeURIComponent(sessionId)}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${Math.floor(sessionMaxAgeMs / 1000)}`,
  ])
}

function clearSessionCookie(response) {
  response.setHeader('Set-Cookie', [
    'univer_session=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0',
  ])
}

async function getRequestUser(request) {
  const sessionId = parseCookies(request).get('univer_session')
  if (!sessionId) return null

  const [sessions, users] = await Promise.all([loadSessions(), loadUsers()])
  const session = sessions.find((entry) => entry.id === sessionId)
  if (!session) return null

  const user = users.find((entry) => entry.id === session.userId && !entry.disabled)
  return user ?? null
}

async function requireUser(request, response) {
  const user = await getRequestUser(request)
  if (!user) {
    sendError(response, 401, 'Unauthorized')
    return null
  }
  return user
}

async function requireAdmin(request, response) {
  const user = await requireUser(request, response)
  if (!user) return null
  if (user.role !== 'admin') {
    sendError(response, 403, 'Forbidden')
    return null
  }
  return user
}

async function handleLogin(request, response) {
  const body = await readJsonBody(request)
  const username = typeof body.username === 'string' ? body.username.trim() : ''
  const password = typeof body.password === 'string' ? body.password : ''
  if (!username || !password) {
    sendError(response, 400, 'Username and password are required')
    return
  }

  const users = await loadUsers()
  const user = users.find((entry) => entry.username === username && !entry.disabled)
  if (!user || !verifyPassword(password, user.passwordHash)) {
    sendError(response, 401, 'Invalid username or password')
    return
  }

  const sessions = await loadSessions()
  const session = {
    id: randomBytes(32).toString('hex'),
    userId: user.id,
    createdAt: new Date().toISOString(),
    expiresAt: new Date(Date.now() + sessionMaxAgeMs).toISOString(),
  }
  await saveSessions([...sessions.filter((entry) => entry.userId !== user.id), session])
  setSessionCookie(response, session.id)
  sendJson(response, 200, { user: publicUser(user) })
}

async function handleLogout(request, response) {
  const sessionId = parseCookies(request).get('univer_session')
  if (sessionId) {
    const sessions = await loadSessions()
    await saveSessions(sessions.filter((entry) => entry.id !== sessionId))
  }
  clearSessionCookie(response)
  response.writeHead(204)
  response.end()
}

async function handleChangeOwnPassword(request, response, currentUser) {
  const body = await readJsonBody(request)
  const password = typeof body.password === 'string' ? body.password : ''
  if (password.length < 6) {
    sendError(response, 400, 'Password must be at least 6 characters.')
    return
  }

  const users = await loadUsers()
  const targetIndex = users.findIndex((user) => user.id === currentUser.id)
  if (targetIndex === -1) {
    sendError(response, 404, 'User not found.')
    return
  }

  const nextUsers = users.map((user, index) =>
    index === targetIndex
      ? { ...user, passwordHash: hashPassword(password), updatedAt: new Date().toISOString() }
      : user,
  )
  await saveUsers(nextUsers)
  sendJson(response, 200, { user: publicUser(nextUsers[targetIndex]) })
}

async function handleListUsers(response) {
  const users = await loadUsers()
  sendJson(response, 200, { users: users.map(publicAdminUser) })
}

async function handleCreateUser(request, response) {
  const body = await readJsonBody(request)
  const username = normalizeUsername(body.username)
  const password = typeof body.password === 'string' ? body.password : ''
  const role = body.isAdmin === true ? 'admin' : normalizeUserRole(body.role)

  if (!isValidUsername(username)) {
    sendError(response, 400, 'Username must be 2-32 characters and only contain letters, numbers, underscore, dot, or hyphen.')
    return
  }

  if (password.length < 6) {
    sendError(response, 400, 'Password must be at least 6 characters.')
    return
  }

  const users = await loadUsers()
  const exists = users.some((user) => user.username.toLowerCase() === username.toLowerCase())
  if (exists) {
    sendError(response, 409, 'Username already exists.')
    return
  }

  const now = new Date().toISOString()
  const user = {
    id: `user-${randomUUID()}`,
    username,
    passwordHash: hashPassword(password),
    role,
    disabled: false,
    createdAt: now,
    updatedAt: now,
  }

  await saveUsers([...users, user])
  await fs.mkdir(path.join(usersRoot, user.id), { recursive: true })
  sendJson(response, 201, { user: publicAdminUser(user) })
}

async function handleUpdateUser(request, response, targetUserId, currentUser) {
  const body = await readJsonBody(request)
  const users = await loadUsers()
  const targetIndex = users.findIndex((user) => user.id === targetUserId)
  if (targetIndex === -1) {
    sendError(response, 404, 'User not found.')
    return
  }

  const target = users[targetIndex]
  const nextUser = { ...target }
  const selfUpdate = currentUser.id === target.id

  if ('role' in body) {
    if (selfUpdate) {
      sendError(response, 400, 'You cannot change your own role.')
      return
    }
    nextUser.role = normalizeUserRole(body.role)
  }

  if ('disabled' in body) {
    if (selfUpdate && body.disabled === true) {
      sendError(response, 400, 'You cannot disable your own account.')
      return
    }
    nextUser.disabled = body.disabled === true
  }

  if ('password' in body) {
    const password = typeof body.password === 'string' ? body.password : ''
    if (password.length < 6) {
      sendError(response, 400, 'Password must be at least 6 characters.')
      return
    }
    nextUser.passwordHash = hashPassword(password)
  }

  const nextUsers = users.map((user, index) =>
    index === targetIndex ? { ...nextUser, updatedAt: new Date().toISOString() } : user,
  )
  const activeAdmins = nextUsers.filter((user) => user.role === 'admin' && !user.disabled)
  if (activeAdmins.length === 0) {
    sendError(response, 400, 'At least one active admin account is required.')
    return
  }

  await saveUsers(nextUsers)

  if (nextUser.disabled) {
    const sessions = await loadSessions()
    await saveSessions(sessions.filter((session) => session.userId !== nextUser.id))
  }

  sendJson(response, 200, { user: publicAdminUser(nextUsers[targetIndex]) })
}

async function readJsonBody(request) {
  const chunks = []
  for await (const chunk of request) {
    chunks.push(chunk)
  }

  if (chunks.length === 0) {
    return {}
  }

  return JSON.parse(Buffer.concat(chunks).toString('utf8'))
}

async function readRawBody(request, maxBytes = 25 * 1024 * 1024) {
  const chunks = []
  let total = 0

  for await (const chunk of request) {
    total += chunk.length
    if (total > maxBytes) {
      throw new Error('File is too large. The MVP importer currently supports files up to 25MB.')
    }
    chunks.push(chunk)
  }

  return Buffer.concat(chunks)
}

function sendBinary(response, statusCode, buffer, headers = {}) {
  response.writeHead(statusCode, headers)
  response.end(buffer)
}

async function ensureDataDirectory() {
  await fs.mkdir(dataRoot, { recursive: true })
}

async function migrateLegacyAIConfigToAdmin() {
  try {
    await fs.access(aiConfigPath)
    return
  } catch {
    // Admin scoped AI config does not exist yet.
  }

  try {
    await fs.mkdir(path.dirname(aiConfigPath), { recursive: true })
    await fs.copyFile(legacyAIConfigPath, aiConfigPath)
    console.log(`[app-server] migrated legacy AI config to admin: ${aiConfigPath}`)
  } catch {
    // Legacy config may not exist; defaults will be used.
  }
}

function normalizeBaseURL(baseURL, provider) {
  const fallback = providerDefaults[provider]?.baseURL ?? ''
  let value = typeof baseURL === 'string' && baseURL.trim() ? baseURL.trim() : fallback
  if (!value) return ''

  value = value
    .replace(/\/+$/, '')
    .replace(/\/chat\/completions$/i, '')
    .replace(/\/models$/i, '')

  if (/\/v\d+$/i.test(value)) {
    return value
  }

  return `${value}/v1`
}

function normalizeModelEntry(entry) {
  return {
    id: typeof entry.id === 'string' && entry.id.trim() ? entry.id : randomUUID(),
    provider: typeof entry.provider === 'string' && entry.provider.trim() ? entry.provider : 'custom',
    model: typeof entry.model === 'string' ? entry.model.trim() : '',
    label: typeof entry.label === 'string' && entry.label.trim()
      ? entry.label.trim()
      : (typeof entry.model === 'string' ? entry.model.trim() : ''),
    apiKey: typeof entry.apiKey === 'string' ? entry.apiKey : '',
    baseURL: normalizeBaseURL(entry.baseURL, entry.provider),
    createdAt: typeof entry.createdAt === 'string' && entry.createdAt ? entry.createdAt : new Date().toISOString(),
    updatedAt: typeof entry.updatedAt === 'string' && entry.updatedAt ? entry.updatedAt : new Date().toISOString(),
  }
}

function normalizeSettings(parsed) {
  if (parsed && Array.isArray(parsed.models)) {
    const models = parsed.models
      .map((entry) => normalizeModelEntry(entry))
      .filter((entry) => entry.model && entry.baseURL)

    const activeModelId = typeof parsed.activeModelId === 'string' ? parsed.activeModelId : ''
    const hasActiveModel = models.some((entry) => entry.id === activeModelId)

    return {
      activeModelId: hasActiveModel ? activeModelId : (models[0]?.id ?? ''),
      models,
    }
  }

  if (parsed && typeof parsed === 'object' && typeof parsed.model === 'string' && parsed.model.trim()) {
    const legacyEntry = normalizeModelEntry({
      id: `legacy-${parsed.provider ?? 'custom'}-${parsed.model.trim()}`,
      provider: parsed.provider,
      model: parsed.model,
      label: parsed.model,
      apiKey: parsed.apiKey,
      baseURL: parsed.baseURL,
    })

    if (legacyEntry.baseURL) {
      return {
        activeModelId: legacyEntry.id,
        models: [legacyEntry],
      }
    }
  }

  return { ...defaultAISettings }
}

async function loadAISettings() {
  try {
    const raw = await fs.readFile(aiConfigPath, 'utf8')
    return normalizeSettings(JSON.parse(raw))
  } catch {
    return { ...defaultAISettings }
  }
}

async function saveAISettings(settings) {
  const normalized = normalizeSettings(settings)
  await ensureDataDirectory()
  await fs.writeFile(aiConfigPath, JSON.stringify(normalized, null, 2), 'utf8')
  return normalized
}

async function clearAISettings() {
  try {
    await fs.rm(aiConfigPath, { force: true })
  } catch {
    // Ignore delete failures for a missing config file.
  }
}

function getModelsEndpoint(baseURL, provider) {
  const normalizedBaseURL = normalizeBaseURL(baseURL, provider)
  if (!normalizedBaseURL) {
    throw new Error('Base URL is required')
  }

  return `${normalizedBaseURL}/models`
}

function getChatCompletionsEndpoint(baseURL, provider = 'custom') {
  const normalizedBaseURL = normalizeBaseURL(baseURL, provider)
  if (!normalizedBaseURL) {
    throw new Error('Base URL is required')
  }

  return `${normalizedBaseURL}/chat/completions`
}

async function handleLLMChatCompletions(request, response) {
  const body = await readJsonBody(request)
  const provider = typeof body.provider === 'string' && body.provider.trim() ? body.provider.trim() : 'custom'
  const apiKey = typeof body.apiKey === 'string' ? body.apiKey.trim() : ''
  const payload = body.payload && typeof body.payload === 'object' ? body.payload : null
  const endpoint = getChatCompletionsEndpoint(body.baseURL, provider)

  if (!payload) {
    sendError(response, 400, 'payload is required')
    return
  }

  const headers = {
    'Content-Type': 'application/json',
  }
  if (apiKey) {
    headers.Authorization = `Bearer ${apiKey}`
  }

  const startedAt = Date.now()
  console.info('[app-server][llm-proxy] request', {
    endpoint,
    model: payload.model,
    messageCount: Array.isArray(payload.messages) ? payload.messages.length : 0,
    toolCount: Array.isArray(payload.tools) ? payload.tools.length : 0,
  })

  const upstreamResponse = await fetch(endpoint, {
    method: 'POST',
    headers,
    body: JSON.stringify(payload),
    signal: AbortSignal.timeout(120000),
  })
  const text = await upstreamResponse.text()
  const contentType = upstreamResponse.headers.get('content-type') || 'application/json; charset=utf-8'

  console.info('[app-server][llm-proxy] response', {
    status: upstreamResponse.status,
    ms: Date.now() - startedAt,
    bytes: text.length,
  })

  response.writeHead(upstreamResponse.status, { 'Content-Type': contentType })
  response.end(text)
}

function getProxyURL() {
  return process.env.APP_HTTP_PROXY?.trim()
    || process.env.HTTPS_PROXY?.trim()
    || process.env.HTTP_PROXY?.trim()
    || ''
}

function psSingleQuote(value) {
  return String(value).replace(/'/g, "''")
}

async function fetchAvailableModelsViaPowerShell({ endpoint, apiKey, proxyURL }) {
  const headersScript = apiKey.trim()
    ? `$headers = @{ Authorization = 'Bearer ${psSingleQuote(apiKey.trim())}' }`
    : '$headers = @{}'
  const proxyScript = proxyURL
    ? ` -Proxy '${psSingleQuote(proxyURL)}'`
    : ''

  const script = [
    `$ProgressPreference = 'SilentlyContinue'`,
    headersScript,
    `$response = Invoke-RestMethod -Uri '${psSingleQuote(endpoint)}' -Method Get -Headers $headers${proxyScript}`,
    `$ids = @()`,
    `if ($response -and $response.data) { $ids = @($response.data | ForEach-Object { $_.id } | Where-Object { $_ }) }`,
    `$ids | ConvertTo-Json -Compress`,
  ].join('; ')

  const encoded = Buffer.from(script, 'utf16le').toString('base64')
  const { stdout } = await execFileAsync('powershell.exe', ['-NoProfile', '-EncodedCommand', encoded], {
    windowsHide: true,
    maxBuffer: 1024 * 1024,
  })

  const parsed = JSON.parse(stdout.trim() || '[]')
  return Array.isArray(parsed)
    ? parsed.filter((item) => typeof item === 'string' && item.trim())
    : (typeof parsed === 'string' && parsed.trim() ? [parsed] : [])
}

async function fetchAvailableModels({ provider, apiKey, baseURL }) {
  const endpoint = getModelsEndpoint(baseURL, provider)
  const proxyURL = getProxyURL()
  const headers = {}
  if (typeof apiKey === 'string' && apiKey.trim()) {
    headers.Authorization = `Bearer ${apiKey.trim()}`
  }

  let models = []
  let lastError = null

  try {
    const response = await fetch(endpoint, { headers })
    if (!response.ok) {
      const text = await response.text()
      throw new Error(`Model discovery failed with status ${response.status}: ${text.slice(0, 200)}`)
    }

    const payload = await response.json()
    models = Array.isArray(payload?.data)
      ? payload.data.map((item) => item?.id).filter((item) => typeof item === 'string' && item.trim())
      : []
  } catch (error) {
    lastError = error

    if (process.platform === 'win32' && proxyURL) {
      models = await fetchAvailableModelsViaPowerShell({
        endpoint,
        apiKey: typeof apiKey === 'string' ? apiKey : '',
        proxyURL,
      })
    } else {
      throw error
    }
  }

  if (!Array.isArray(models)) {
    throw lastError ?? new Error('Model discovery failed')
  }

  return {
    baseURL: normalizeBaseURL(baseURL, provider),
    models: Array.from(new Set(models)).sort((a, b) => a.localeCompare(b)),
  }
}

async function addModelsToSettings(body) {
  const provider = typeof body.provider === 'string' && body.provider.trim() ? body.provider.trim() : 'custom'
  const apiKey = typeof body.apiKey === 'string' ? body.apiKey.trim() : ''
  const baseURL = normalizeBaseURL(body.baseURL, provider)
  const models = Array.isArray(body.models)
    ? body.models.map((item) => (typeof item === 'string' ? item.trim() : '')).filter(Boolean)
    : []

  if (!baseURL) {
    throw new Error('Base URL is required')
  }

  if (models.length === 0) {
    throw new Error('At least one model is required')
  }

  const settings = await loadAISettings()
  const now = new Date().toISOString()
  const nextModels = [...settings.models]
  let nextActiveModelId = settings.activeModelId

  for (const model of models) {
    const existing = nextModels.find((entry) =>
      entry.provider === provider &&
      entry.baseURL === baseURL &&
      entry.model === model,
    )

    if (existing) {
      existing.apiKey = apiKey
      existing.label = model
      existing.updatedAt = now
      nextActiveModelId = existing.id
      continue
    }

    const created = normalizeModelEntry({
      id: randomUUID(),
      provider,
      model,
      label: model,
      apiKey,
      baseURL,
      createdAt: now,
      updatedAt: now,
    })

    nextModels.push(created)
    nextActiveModelId = created.id
  }

  return saveAISettings({
    activeModelId: nextActiveModelId || settings.activeModelId,
    models: nextModels,
  })
}

async function setActiveModel(modelId) {
  const settings = await loadAISettings()
  const exists = settings.models.some((entry) => entry.id === modelId)
  if (!exists) {
    throw new Error('Model not found')
  }

  return saveAISettings({
    ...settings,
    activeModelId: modelId,
  })
}

async function deleteModel(modelId) {
  const settings = await loadAISettings()
  const nextModels = settings.models.filter((entry) => entry.id !== modelId)
  const nextActiveModelId = settings.activeModelId === modelId
    ? (nextModels[0]?.id ?? '')
    : settings.activeModelId

  return saveAISettings({
    activeModelId: nextActiveModelId,
    models: nextModels,
  })
}

function decodeFileName(value) {
  if (typeof value !== 'string' || !value.trim()) return 'Imported.xlsx'

  try {
    return decodeURIComponent(value)
  } catch {
    return value
  }
}

function stripWorkbookExtension(fileName) {
  return fileName.replace(/\.(xlsx|xlsm|xls)$/i, '') || 'Imported Workbook'
}

function sanitizeDownloadFileName(fileName) {
  const cleaned = String(fileName || 'workbook.xlsx')
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_')
    .trim()
  const withExtension = /\.(xlsx)$/i.test(cleaned) ? cleaned : `${cleaned || 'workbook'}.xlsx`
  return withExtension
}

function argbToHex(value) {
  if (!value || typeof value !== 'string') return null
  const normalized = value.replace(/^#/, '')
  if (/^[0-9a-f]{8}$/i.test(normalized)) {
    const alpha = normalized.slice(0, 2).toUpperCase()
    const rgb = normalized.slice(2).toUpperCase()
    if (alpha === '00' && rgb === '000000') return null
    return `#${rgb}`
  }
  if (/^[0-9a-f]{6}$/i.test(normalized)) return `#${normalized.toUpperCase()}`
  return null
}

const defaultThemeColors = [
  'FFFFFF',
  '000000',
  'EEECE1',
  '1F497D',
  '4F81BD',
  'C0504D',
  '9BBB59',
  '8064A2',
  '4BACC6',
  'F79646',
  '0000FF',
  '800080',
]

const indexedColors = {
  0: '000000',
  1: 'FFFFFF',
  2: 'FF0000',
  3: '00FF00',
  4: '0000FF',
  5: 'FFFF00',
  6: 'FF00FF',
  7: '00FFFF',
  8: '000000',
  9: 'FFFFFF',
  10: 'FF0000',
  11: '00FF00',
  12: '0000FF',
  13: 'FFFF00',
  14: 'FF00FF',
  15: '00FFFF',
  16: '800000',
  17: '008000',
  18: '000080',
  19: '808000',
  20: '800080',
  21: '008080',
  22: 'C0C0C0',
  23: '808080',
  24: '9999FF',
  25: '993366',
  26: 'FFFFCC',
  27: 'CCFFFF',
  28: '660066',
  29: 'FF8080',
  30: '0066CC',
  31: 'CCCCFF',
  32: '000080',
  33: 'FF00FF',
  34: 'FFFF00',
  35: '00FFFF',
  36: '800080',
  37: '800000',
  38: '008080',
  39: '0000FF',
  40: '00CCFF',
  41: 'CCFFFF',
  42: 'CCFFCC',
  43: 'FFFF99',
  44: '99CCFF',
  45: 'FF99CC',
  46: 'CC99FF',
  47: 'FFCC99',
  48: '3366FF',
  49: '33CCCC',
  50: '99CC00',
  51: 'FFCC00',
  52: 'FF9900',
  53: 'FF6600',
  54: '666699',
  55: '969696',
  56: '003366',
  57: '339966',
  58: '003300',
  59: '333300',
  60: '993300',
  61: '993366',
  62: '333399',
  63: '333333',
  64: null,
  65: null,
}

const cssNamedColors = {
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

function extractWorkbookThemeColors(workbook) {
  const themeXml = getWorkbookThemeXml(workbook)
  if (!themeXml) return defaultThemeColors

  const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']
  const colors = order.map((name) => {
    const blockMatch = new RegExp(`<a:${name}\\b[^>]*>([\\s\\S]*?)<\\/a:${name}>`, 'i').exec(themeXml)
    const block = blockMatch?.[1] ?? ''
    const srgbMatch = /<a:srgbClr\b[^>]*\bval="([0-9a-f]{6})"/i.exec(block)
    if (srgbMatch) return srgbMatch[1].toUpperCase()
    const sysMatch = /<a:sysClr\b[^>]*\blastClr="([0-9a-f]{6})"/i.exec(block)
    if (sysMatch) return sysMatch[1].toUpperCase()
    return null
  })

  return colors.every(Boolean) ? colors : defaultThemeColors
}

function getWorkbookThemeXml(workbook) {
  const themes = workbook?.model?.themes ?? workbook?._themes
  if (!themes) return null
  if (typeof themes === 'string') return themes
  if (Array.isArray(themes)) return themes.find((theme) => typeof theme === 'string' && theme.includes('<a:theme')) ?? null
  if (typeof themes === 'object') {
    return Object.values(themes).find((theme) => typeof theme === 'string' && theme.includes('<a:theme')) ?? null
  }
  return null
}

function excelColorToHex(color, themeColors = defaultThemeColors) {
  if (!color || typeof color !== 'object') return null
  const argbColor = argbToHex(color.argb)
  if (argbColor) return applyTint(argbColor, color.tint)

  const themeIndex = Number(color.theme)
  if (Number.isInteger(themeIndex)) {
    const themeColor = themeColors[themeIndex]
    if (themeColor) return applyTint(`#${themeColor}`, color.tint)
  }

  const indexedIndex = Number(color.indexed)
  if (Number.isInteger(indexedIndex)) {
    const indexedColor = indexedColors[indexedIndex]
    if (indexedColor) return applyTint(`#${indexedColor}`, color.tint)
  }

  return null
}

function applyTint(hexColor, tint) {
  const tintValue = Number(tint)
  if (!Number.isFinite(tintValue) || tintValue === 0) return hexColor
  const rgb = hexToRgb(hexColor)
  if (!rgb) return hexColor

  const hsl = rgbToHsl(rgb[0], rgb[1], rgb[2])
  hsl[2] = tintValue < 0
    ? hsl[2] * (1 + tintValue)
    : hsl[2] * (1 - tintValue) + tintValue
  const adjusted = hslToRgb(hsl[0], hsl[1], Math.max(0, Math.min(1, hsl[2])))

  return `#${adjusted.map((channel) => channel.toString(16).padStart(2, '0')).join('').toUpperCase()}`
}

function hexToRgb(hexColor) {
  const normalized = String(hexColor).replace(/^#/, '')
  if (!/^[0-9a-f]{6}$/i.test(normalized)) return null
  return [
    Number.parseInt(normalized.slice(0, 2), 16),
    Number.parseInt(normalized.slice(2, 4), 16),
    Number.parseInt(normalized.slice(4, 6), 16),
  ]
}

function rgbToHsl(red, green, blue) {
  const r = red / 255
  const g = green / 255
  const b = blue / 255
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const lightness = (max + min) / 2

  if (max === min) return [0, 0, lightness]

  const delta = max - min
  const saturation = lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min)
  let hue = 0

  if (max === r) hue = (g - b) / delta + (g < b ? 6 : 0)
  else if (max === g) hue = (b - r) / delta + 2
  else hue = (r - g) / delta + 4

  return [hue / 6, saturation, lightness]
}

function hslToRgb(hue, saturation, lightness) {
  if (saturation === 0) {
    const value = Math.round(lightness * 255)
    return [value, value, value]
  }

  const q = lightness < 0.5 ? lightness * (1 + saturation) : lightness + saturation - lightness * saturation
  const p = 2 * lightness - q
  return [hue + 1 / 3, hue, hue - 1 / 3].map((channel) => {
    let t = channel
    if (t < 0) t += 1
    if (t > 1) t -= 1
    if (t < 1 / 6) return Math.round((p + (q - p) * 6 * t) * 255)
    if (t < 1 / 2) return Math.round(q * 255)
    if (t < 2 / 3) return Math.round((p + (q - p) * (2 / 3 - t) * 6) * 255)
    return Math.round(p * 255)
  })
}

function excelFillToBackground(fill, themeColors) {
  if (!fill || typeof fill !== 'object') return null
  if (fill.type === 'pattern') return excelPatternFillToBackground(fill, themeColors)

  return excelColorToHex(fill.fgColor, themeColors)
    ?? excelColorToHex(fill.bgColor, themeColors)
    ?? getGradientFillFallback(fill, themeColors)
}

function excelPatternFillToBackground(fill, themeColors) {
  const pattern = String(fill.pattern || '').trim()
  if (!pattern || pattern === 'none') return null

  const foreground = excelColorToHex(fill.fgColor, themeColors)
  const background = excelColorToHex(fill.bgColor, themeColors)

  if (pattern === 'solid') {
    if (isNearBlack(foreground) && isVisibleNonBlackColor(background)) {
      return background
    }
    return foreground ?? background
  }

  // Excel uses gray125/gray0625 as pattern placeholders in many styles.
  // They should not become a visible black background in Univer.
  if (pattern === 'gray125' || pattern === 'gray0625') {
    return background
  }

  const foregroundWeight = getPatternForegroundWeight(pattern)
  if (foregroundWeight === null) {
    return foreground ?? background
  }

  return blendHexColors(background ?? '#FFFFFF', foreground ?? '#000000', foregroundWeight)
}

function isNearBlack(hexColor) {
  const rgb = hexToRgb(hexColor)
  if (!rgb) return false
  return rgb.every((channel) => channel <= 12)
}

function isVisibleNonBlackColor(hexColor) {
  const rgb = hexToRgb(hexColor)
  if (!rgb) return false
  return rgb.some((channel) => channel > 12)
}

function getPatternForegroundWeight(pattern) {
  const normalized = pattern.toLowerCase()
  const weights = {
    darkgray: 0.75,
    mediumgray: 0.5,
    lightgray: 0.25,
    darkhorizontal: 0.55,
    darkvertical: 0.55,
    darkdown: 0.55,
    darkup: 0.55,
    darkgrid: 0.6,
    darktrellis: 0.6,
    lighthorizontal: 0.25,
    lightvertical: 0.25,
    lightdown: 0.25,
    lightup: 0.25,
    lightgrid: 0.3,
    lighttrellis: 0.3,
  }
  return weights[normalized] ?? null
}

function blendHexColors(backgroundHex, foregroundHex, foregroundWeight) {
  const background = hexToRgb(backgroundHex)
  const foreground = hexToRgb(foregroundHex)
  if (!background || !foreground) return foregroundHex

  const weight = Math.max(0, Math.min(1, foregroundWeight))
  const blended = background.map((channel, index) =>
    Math.round(channel * (1 - weight) + foreground[index] * weight),
  )
  return `#${blended.map((channel) => channel.toString(16).padStart(2, '0')).join('').toUpperCase()}`
}

function getGradientFillFallback(fill, themeColors) {
  if (!Array.isArray(fill.stops) || fill.stops.length === 0) return null
  const stop = fill.stops.find((item) => item?.color) ?? fill.stops[0]
  return excelColorToHex(stop?.color, themeColors)
}

function excelDateToSerialNumber(value, date1904 = false) {
  const epoch = Date.UTC(date1904 ? 1904 : 1899, date1904 ? 0 : 11, date1904 ? 1 : 30)
  return (value.getTime() - epoch) / 86400000
}

function excelCellValueToPersistedValue(value, date1904 = false) {
  if (value === null || value === undefined) return null
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value
  if (value instanceof Date) return excelDateToSerialNumber(value, date1904)
  if (typeof value === 'object') {
    if ('result' in value) return excelCellValueToPersistedValue(value.result, date1904)
    if ('text' in value) return String(value.text)
    if ('richText' in value && Array.isArray(value.richText)) {
      return value.richText.map((part) => part?.text ?? '').join('')
    }
    if ('hyperlink' in value && 'text' in value) return String(value.text)
  }
  return String(value)
}

function excelCellFormulaToPersistedFormula(value) {
  if (!value || typeof value !== 'object' || !('formula' in value)) return ''
  const formula = typeof value.formula === 'string' ? value.formula.trim() : ''
  if (!formula) return ''
  return formula.startsWith('=') ? formula : `=${formula}`
}

function excelBorderToPersistedBorder(border, themeColors) {
  if (!border || typeof border !== 'object') return null
  const result = {}

  for (const side of ['top', 'right', 'bottom', 'left']) {
    const edge = border[side]
    if (edge?.style) {
      result[side] = {
        style: edge.style,
        color: excelColorToHex(edge.color, themeColors),
      }
    }
  }

  return Object.keys(result).length > 0 ? result : null
}

function excelCellToPersistedStyle(cell, themeColors) {
  const style = {}
  const font = cell.font ?? null
  const fill = cell.fill ?? null
  const alignment = cell.alignment ?? null
  const border = excelBorderToPersistedBorder(cell.border, themeColors)

  if (typeof font?.name === 'string' && font.name.trim()) style.fontFamily = font.name.trim()
  if (font?.bold) style.fontWeight = 'bold'
  if (font?.italic) style.fontStyle = 'italic'
  if (font?.underline) style.underline = true
  if (font?.strike) style.strikethrough = true
  if (typeof font?.size === 'number') style.fontSize = font.size
  const fontColor = excelColorToHex(font?.color, themeColors)
  if (fontColor) style.fontColor = fontColor

  const background = excelFillToBackground(fill, themeColors)
  if (background) style.background = background

  if (alignment?.horizontal) style.horizontalAlignment = alignment.horizontal
  if (alignment?.vertical) style.verticalAlignment = alignment.vertical
  if (alignment?.wrapText) style.wrap = true
  if (typeof alignment?.textRotation === 'number') style.textRotation = alignment.textRotation
  if (cell.numFmt) style.numFmt = cell.numFmt
  if (border) style.border = border

  return Object.keys(style).length > 0 ? style : null
}

function logSuspiciousImportedColor({ debugState, fileName, sheetName, cell, style, themeColors }) {
  if (!style?.background || !isNearBlack(style.background)) return
  if (!debugState || debugState.count >= debugState.limit) return

  debugState.count += 1
  console.warn('[xlsx-color-debug]', JSON.stringify({
    fileName,
    sheetName,
    address: cell.address,
    value: excelCellValueToPersistedValue(cell.value),
    computedStyle: style,
    fill: toLoggableObject(cell.fill),
    font: toLoggableObject(cell.font),
    border: toLoggableObject(cell.border),
    numFmt: cell.numFmt || null,
    styleId: cell.styleId ?? null,
    rawStyle: toLoggableObject(cell.style),
    themeColors,
    note: 'Computed background is near black. Copy this block to diagnose Excel color parsing.',
  }))

  if (debugState.count === debugState.limit) {
    console.warn('[xlsx-color-debug]', JSON.stringify({
      fileName,
      message: `Reached color debug log limit (${debugState.limit}). Set APP_XLSX_COLOR_DEBUG_LIMIT to increase it.`,
    }))
  }
}

function getImportDiagnostics(request) {
  const enabled = isTruthyHeader(request.headers['x-import-diagnostics'])
    || process.env.APP_XLSX_IMPORT_DIAGNOSTICS === '1'
    || process.env.APP_XLSX_COLOR_DEBUG === '1'
  const limit = Number.isFinite(maxColorDebugLogs) && maxColorDebugLogs > 0 ? Math.floor(maxColorDebugLogs) : 80
  return {
    enabled,
    colorDebugState: enabled ? { count: 0, limit } : null,
  }
}

function isTruthyHeader(value) {
  const raw = Array.isArray(value) ? value[0] : value
  if (typeof raw !== 'string') return false
  return ['1', 'true', 'yes', 'on'].includes(raw.trim().toLowerCase())
}

function toLoggableObject(value) {
  if (value === null || value === undefined) return null
  try {
    return JSON.parse(JSON.stringify(value))
  } catch {
    return String(value)
  }
}

function countStyledCells(styles) {
  let total = 0
  for (const row of styles) {
    for (const style of row) {
      if (style) total += 1
    }
  }
  return total
}

function collectHiddenRowRanges(worksheet, maxRow) {
  const hiddenRows = []
  let rangeStart = -1

  for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
    const hidden = worksheet.getRow(rowNumber)?.hidden === true
    if (hidden && rangeStart === -1) {
      rangeStart = rowNumber - 1
    }
    if ((!hidden || rowNumber === maxRow) && rangeStart !== -1) {
      hiddenRows.push({
        start: rangeStart,
        end: hidden ? rowNumber - 1 : rowNumber - 2,
      })
      rangeStart = -1
    }
  }

  return hiddenRows
}

function collectHiddenColumnRanges(worksheet, maxColumn) {
  const hiddenColumns = []
  let rangeStart = -1

  for (let columnNumber = 1; columnNumber <= maxColumn; columnNumber += 1) {
    const hidden = worksheet.getColumn(columnNumber)?.hidden === true
    if (hidden && rangeStart === -1) {
      rangeStart = columnNumber - 1
    }
    if ((!hidden || columnNumber === maxColumn) && rangeStart !== -1) {
      hiddenColumns.push({
        start: rangeStart,
        end: hidden ? columnNumber - 1 : columnNumber - 2,
      })
      rangeStart = -1
    }
  }

  return hiddenColumns
}

function excelWorksheetViewToPersistedView(worksheet, themeColors = defaultThemeColors) {
  const view = Array.isArray(worksheet.views) ? worksheet.views[0] : null
  const result = {}
  const tabColor = excelColorToHex(worksheet.properties?.tabColor, themeColors)

  if (view?.state === 'frozen') {
    if (Number(view.ySplit) > 0) result.frozenRows = Math.floor(Number(view.ySplit))
    if (Number(view.xSplit) > 0) result.frozenColumns = Math.floor(Number(view.xSplit))
  }

  if (view && view.showGridLines === false) {
    result.hiddenGridlines = true
  }

  if (tabColor) {
    result.tabColor = tabColor
  }

  return result
}

function getWorksheetMergedRanges(worksheet) {
  const modelMerges = Array.isArray(worksheet.model?.merges) ? worksheet.model.merges : []
  if (modelMerges.length > 0) return modelMerges

  const privateMerges = worksheet._merges && typeof worksheet._merges === 'object'
    ? Object.keys(worksheet._merges)
    : []
  return privateMerges
}

function columnWidthToPixels(width) {
  return Math.max(24, Math.round((Number(width) || 8.43) * 7 + 5))
}

function columnPixelsToWidth(width) {
  return Math.max(1, Math.round(((Number(width) || 64) - 5) / 7 * 100) / 100)
}

function rowHeightToPixels(height) {
  return Math.max(12, Math.round((Number(height) || 15) * 4 / 3))
}

function rowPixelsToHeight(height) {
  return Math.max(1, Math.round((Number(height) || 20) * 0.75 * 100) / 100)
}

function columnLetterToNumber(columnLetters) {
  let result = 0
  for (const character of String(columnLetters).toUpperCase()) {
    result = result * 26 + character.charCodeAt(0) - 64
  }
  return result
}

function getRangeBounds(rangeA1) {
  const normalized = String(rangeA1).replace(/\$/g, '')
  const match = normalized.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i)
  if (!match) return null

  const startColumn = columnLetterToNumber(match[1])
  const startRow = Number(match[2])
  const endColumn = match[3] ? columnLetterToNumber(match[3]) : startColumn
  const endRow = match[4] ? Number(match[4]) : startRow

  return {
    startRow,
    startColumn,
    endRow,
    endColumn,
  }
}

function getUsedBounds(worksheet, mergedRanges) {
  let maxRow = 0
  let maxColumn = 0

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
      if (cell.value === null || cell.value === undefined || cell.value === '') return
      maxRow = Math.max(maxRow, rowNumber)
      maxColumn = Math.max(maxColumn, columnNumber)
    })
  })

  for (const mergedRange of mergedRanges) {
    const bounds = getRangeBounds(mergedRange)
    if (!bounds) continue
    maxRow = Math.max(maxRow, bounds.endRow)
    maxColumn = Math.max(maxColumn, bounds.endColumn)
  }

  return { maxRow, maxColumn }
}

function assertImportSize(sheetName, maxRow, maxColumn) {
  const cellCount = maxRow * maxColumn
  if (cellCount > maxImportCells) {
    throw new Error(`工作表「${sheetName}」使用范围过大：${maxRow} 行 x ${maxColumn} 列。当前 MVP 限制 ${maxImportCells} 个单元格，请先在 Excel 中删除无效空白行列或拆分文件。`)
  }
}

function normalizePersistedWorkbook(workbook) {
  const sheets = Array.isArray(workbook?.sheets)
    ? workbook.sheets.filter((sheet) => sheet && typeof sheet.name === 'string')
    : []

  return {
    version: 'grid-v1',
    name: typeof workbook?.name === 'string' && workbook.name.trim() ? workbook.name.trim() : 'Workbook',
    fileName: typeof workbook?.fileName === 'string' && workbook.fileName.trim()
      ? sanitizeDownloadFileName(workbook.fileName.trim())
      : 'workbook.xlsx',
    activeSheetName: typeof workbook?.activeSheetName === 'string' && workbook.activeSheetName.trim()
      ? workbook.activeSheetName.trim()
      : (sheets[0]?.name ?? 'Sheet1'),
    sheets: sheets.map((sheet) => ({
      name: sheet.name,
      values: normalizeMatrix(sheet.values),
      formulas: normalizeMatrix(sheet.formulas),
      styles: normalizeMatrix(sheet.styles),
      columnWidths: sheet.columnWidths && typeof sheet.columnWidths === 'object' ? sheet.columnWidths : {},
      rowHeights: sheet.rowHeights && typeof sheet.rowHeights === 'object' ? sheet.rowHeights : {},
      mergedRanges: Array.isArray(sheet.mergedRanges) ? sheet.mergedRanges.filter((item) => typeof item === 'string') : [],
      hiddenRows: normalizeIndexRanges(sheet.hiddenRows),
      hiddenColumns: normalizeIndexRanges(sheet.hiddenColumns),
      sheetView: normalizeSheetView(sheet.sheetView),
    })),
  }
}

function normalizeMatrix(value) {
  return Array.isArray(value)
    ? value.map((row) => (Array.isArray(row) ? row : []))
    : []
}

function normalizeIndexRanges(value) {
  if (!Array.isArray(value)) return []
  return value
    .map((range) => ({
      start: Number(range?.start),
      end: Number(range?.end),
    }))
    .filter((range) =>
      Number.isInteger(range.start) &&
      Number.isInteger(range.end) &&
      range.start >= 0 &&
      range.end >= range.start,
    )
}

function normalizeSheetView(value) {
  if (!value || typeof value !== 'object') return {}
  const result = {}
  if (Number.isInteger(value.frozenRows) && value.frozenRows > 0) result.frozenRows = value.frozenRows
  if (Number.isInteger(value.frozenColumns) && value.frozenColumns > 0) result.frozenColumns = value.frozenColumns
  if (typeof value.hiddenGridlines === 'boolean') result.hiddenGridlines = value.hiddenGridlines
  if (typeof value.gridlinesColor === 'string' && value.gridlinesColor.trim()) result.gridlinesColor = value.gridlinesColor.trim()
  if (typeof value.tabColor === 'string' && value.tabColor.trim()) result.tabColor = value.tabColor.trim()
  return result
}

function cssHexToArgb(value) {
  const normalized = normalizeCssHexColor(value)
  return normalized ? `FF${normalized.replace(/^#/, '')}` : null
}

function normalizeCssHexColor(value) {
  if (typeof value !== 'string') return null
  const trimmed = value.trim()
  const named = cssNamedColors[trimmed.toLowerCase()]
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
  return null
}

function excelBorderStyle(style) {
  if (typeof style === 'number') {
    return {
      1: 'thin',
      2: 'hair',
      3: 'dotted',
      4: 'dashed',
      5: 'dashDot',
      6: 'dashDotDot',
      7: 'double',
      8: 'medium',
      9: 'mediumDashed',
      10: 'mediumDashDot',
      11: 'mediumDashDotDot',
      12: 'slantDashDot',
      13: 'thick',
    }[style] ?? 'thin'
  }
  if (!style || typeof style !== 'string') return 'thin'
  return style
}

function applyPersistedStyleToExcelCell(cell, style) {
  if (!style || typeof style !== 'object') return

  const font = {}
  if (typeof style.fontFamily === 'string') font.name = style.fontFamily
  if (style.fontWeight === 'bold') font.bold = true
  if (style.fontStyle === 'italic') font.italic = true
  if (style.underline === true || style.fontLine === 'underline') font.underline = true
  if (style.strikethrough === true || style.fontLine === 'line-through') font.strike = true
  if (typeof style.fontSize === 'number') font.size = style.fontSize
  const fontColorArgb = cssHexToArgb(style.fontColor)
  if (fontColorArgb) font.color = { argb: fontColorArgb }
  if (Object.keys(font).length > 0) cell.font = font

  const backgroundArgb = cssHexToArgb(style.background)
  if (backgroundArgb) {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: backgroundArgb },
    }
  }

  const alignment = {}
  if (typeof style.horizontalAlignment === 'string') alignment.horizontal = style.horizontalAlignment
  if (typeof style.verticalAlignment === 'string') alignment.vertical = style.verticalAlignment
  if (style.wrap) alignment.wrapText = true
  if (typeof style.textRotation === 'number') alignment.textRotation = style.textRotation
  if (Object.keys(alignment).length > 0) cell.alignment = alignment

  if (typeof style.numFmt === 'string') cell.numFmt = style.numFmt

  if (style.border && typeof style.border === 'object') {
    const border = {}
    for (const side of ['top', 'right', 'bottom', 'left']) {
      const edge = style.border[side]
      if (edge && typeof edge === 'object') {
        border[side] = {
          style: excelBorderStyle(edge.style),
          color: cssHexToArgb(edge.color) ? { argb: cssHexToArgb(edge.color) } : undefined,
        }
      }
    }
    if (Object.keys(border).length > 0) cell.border = border
  }
}

async function importXlsxWorkbook(request) {
  const fileName = decodeFileName(request.headers['x-file-name'])
  const buffer = await readRawBody(request)
  const excelWorkbook = new ExcelJS.Workbook()
  await excelWorkbook.xlsx.load(buffer)
  const date1904 = excelWorkbook.properties?.date1904 === true
  const themeColors = extractWorkbookThemeColors(excelWorkbook)
  const diagnostics = getImportDiagnostics(request)
  const colorDebugState = diagnostics.colorDebugState

  if (diagnostics.enabled) {
    console.info('[xlsx-import-diagnostics]', JSON.stringify({
      fileName,
      worksheets: excelWorkbook.worksheets.map((worksheet) => worksheet.name),
      date1904,
      themeColors,
      colorDebugLimit: colorDebugState?.limit ?? 0,
    }))
  }

  const sheets = excelWorkbook.worksheets.map((worksheet) => {
    const mergedRanges = getWorksheetMergedRanges(worksheet)
    const { maxRow, maxColumn } = getUsedBounds(worksheet, mergedRanges)
    assertImportSize(worksheet.name, maxRow, maxColumn)

    const values = Array.from({ length: maxRow }, () => Array.from({ length: maxColumn }, () => null))
    const formulas = Array.from({ length: maxRow }, () => Array.from({ length: maxColumn }, () => ''))
    const styles = Array.from({ length: maxRow }, () => Array.from({ length: maxColumn }, () => null))
    const columnWidths = {}
    const rowHeights = {}
    const sheetView = excelWorksheetViewToPersistedView(worksheet, themeColors)

    const worksheetColumns = Array.isArray(worksheet.columns) ? worksheet.columns : []
    worksheetColumns.forEach((column, index) => {
      if (index < maxColumn && column?.width) {
        columnWidths[String(index)] = columnWidthToPixels(column.width)
      }
    })

    for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
      const row = worksheet.getRow(rowNumber)
      if (row.height) {
        rowHeights[String(rowNumber - 1)] = rowHeightToPixels(row.height)
      }
    }

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > maxRow) return

      row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
        if (columnNumber > maxColumn) return

        const rowIndex = rowNumber - 1
        const columnIndex = columnNumber - 1
        if (cell.isMerged && cell.master && cell.master.address !== cell.address) return

        values[rowIndex][columnIndex] = excelCellValueToPersistedValue(cell.value, date1904)
        formulas[rowIndex][columnIndex] = excelCellFormulaToPersistedFormula(cell.value)
      })
    })

    for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
      for (let columnNumber = 1; columnNumber <= maxColumn; columnNumber += 1) {
        const cell = worksheet.getCell(rowNumber, columnNumber)
        const style = excelCellToPersistedStyle(cell, themeColors)
        if (style) {
          styles[rowNumber - 1][columnNumber - 1] = style
          logSuspiciousImportedColor({
            debugState: colorDebugState,
            fileName,
            sheetName: worksheet.name,
            cell,
            style,
            themeColors,
          })
        }
      }
    }

    const hiddenRows = collectHiddenRowRanges(worksheet, maxRow)
    const hiddenColumns = collectHiddenColumnRanges(worksheet, maxColumn)

    return {
      name: worksheet.name,
      values,
      formulas,
      styles,
      columnWidths,
      rowHeights,
      mergedRanges,
      hiddenRows,
      hiddenColumns,
      sheetView,
      importStats: {
        rows: maxRow,
        columns: maxColumn,
        cells: maxRow * maxColumn,
        formulas: formulas.reduce((total, row) => total + row.filter(Boolean).length, 0),
        styledCells: countStyledCells(styles),
        mergedRanges: mergedRanges.length,
        hiddenRows: hiddenRows.reduce((total, range) => total + range.end - range.start + 1, 0),
        hiddenColumns: hiddenColumns.reduce((total, range) => total + range.end - range.start + 1, 0),
      },
    }
  })

  const importSummary = {
    sheets: sheets.length,
    rows: sheets.reduce((total, sheet) => Math.max(total, sheet.importStats?.rows ?? 0), 0),
    columns: sheets.reduce((total, sheet) => Math.max(total, sheet.importStats?.columns ?? 0), 0),
    cells: sheets.reduce((total, sheet) => total + (sheet.importStats?.cells ?? 0), 0),
    formulas: sheets.reduce((total, sheet) => total + (sheet.importStats?.formulas ?? 0), 0),
    styledCells: sheets.reduce((total, sheet) => total + (sheet.importStats?.styledCells ?? 0), 0),
    mergedRanges: sheets.reduce((total, sheet) => total + (sheet.importStats?.mergedRanges ?? 0), 0),
    hiddenRows: sheets.reduce((total, sheet) => total + (sheet.importStats?.hiddenRows ?? 0), 0),
    hiddenColumns: sheets.reduce((total, sheet) => total + (sheet.importStats?.hiddenColumns ?? 0), 0),
  }

  if (diagnostics.enabled) {
    console.info('[xlsx-import-diagnostics]', JSON.stringify({
      fileName,
      importSummary,
      suspiciousColorLogs: colorDebugState?.count ?? 0,
    }))
  }

  return {
    version: 'grid-v1',
    name: stripWorkbookExtension(fileName),
    fileName: sanitizeDownloadFileName(fileName),
    activeSheetName: sheets[0]?.name ?? 'Sheet1',
    importSummary,
    sheets: sheets.length > 0
      ? sheets
      : [{ name: 'Sheet1', values: [], formulas: [], styles: [], columnWidths: {}, rowHeights: {}, mergedRanges: [] }],
  }
}

async function exportXlsxWorkbook(body) {
  const persisted = normalizePersistedWorkbook(body)
  const excelWorkbook = new ExcelJS.Workbook()
  excelWorkbook.creator = 'my-univer-agent'
  excelWorkbook.created = new Date()
  excelWorkbook.modified = new Date()

  for (const sheetData of persisted.sheets) {
    const worksheet = excelWorkbook.addWorksheet(sheetData.name || 'Sheet')
    const sheetView = sheetData.sheetView ?? {}
    const tabColorArgb = cssHexToArgb(sheetView.tabColor)
    if (tabColorArgb) {
      worksheet.properties.tabColor = { argb: tabColorArgb }
    }

    if (sheetView.hiddenGridlines || sheetView.frozenRows || sheetView.frozenColumns) {
      worksheet.views = [{
        state: sheetView.frozenRows || sheetView.frozenColumns ? 'frozen' : 'normal',
        xSplit: sheetView.frozenColumns || 0,
        ySplit: sheetView.frozenRows || 0,
        showGridLines: sheetView.hiddenGridlines ? false : undefined,
      }]
    }

    const rowCount = Math.max(sheetData.values.length, sheetData.formulas.length, sheetData.styles.length)
    const columnCount = Math.max(
      0,
      ...sheetData.values.map((row) => row.length),
      ...sheetData.formulas.map((row) => row.length),
      ...sheetData.styles.map((row) => row.length),
    )

    for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
      const row = worksheet.getRow(rowIndex + 1)
      const rowHeight = Number(sheetData.rowHeights[String(rowIndex)])
      if (Number.isFinite(rowHeight) && rowHeight > 0) {
        row.height = rowPixelsToHeight(rowHeight)
      }

      for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
        const cell = row.getCell(columnIndex + 1)
        const formula = sheetData.formulas[rowIndex]?.[columnIndex]
        const value = sheetData.values[rowIndex]?.[columnIndex] ?? null

        if (formula) {
          cell.value = { formula: formula.replace(/^=/, ''), result: value }
        } else {
          cell.value = value
        }

        applyPersistedStyleToExcelCell(cell, sheetData.styles[rowIndex]?.[columnIndex] ?? null)
      }
    }

    Object.entries(sheetData.columnWidths).forEach(([columnIndex, width]) => {
      const index = Number(columnIndex)
      if (Number.isInteger(index) && index >= 0 && Number(width) > 0) {
        worksheet.getColumn(index + 1).width = columnPixelsToWidth(Number(width))
      }
    })

    for (const range of sheetData.hiddenRows) {
      for (let rowIndex = range.start; rowIndex <= range.end; rowIndex += 1) {
        worksheet.getRow(rowIndex + 1).hidden = true
      }
    }

    for (const range of sheetData.hiddenColumns) {
      for (let columnIndex = range.start; columnIndex <= range.end; columnIndex += 1) {
        worksheet.getColumn(columnIndex + 1).hidden = true
      }
    }

    for (const mergedRange of sheetData.mergedRanges) {
      try {
        worksheet.mergeCells(mergedRange)
      } catch {
        // Keep export best-effort; invalid merge ranges should not block the whole file.
      }
    }
  }

  if (excelWorkbook.worksheets.length === 0) {
    excelWorkbook.addWorksheet('Sheet1')
  }

  return {
    fileName: persisted.fileName,
    buffer: await excelWorkbook.xlsx.writeBuffer(),
  }
}

async function handleApi(request, response, url) {
  if (request.method === 'OPTIONS') {
    response.writeHead(204)
    response.end()
    return
  }

  if (url.pathname === '/api/auth/me' && request.method === 'GET') {
    const user = await getRequestUser(request)
    sendJson(response, 200, { authenticated: !!user, user: publicUser(user) })
    return
  }

  if (url.pathname === '/api/auth/login' && request.method === 'POST') {
    await handleLogin(request, response)
    return
  }

  if (url.pathname === '/api/auth/logout' && request.method === 'POST') {
    await handleLogout(request, response)
    return
  }

  const currentUser = await requireUser(request, response)
  if (!currentUser) return

  if (url.pathname === '/api/auth/password' && request.method === 'PUT') {
    await handleChangeOwnPassword(request, response, currentUser)
    return
  }

  if (url.pathname === '/api/admin/users' && request.method === 'GET') {
    if (!await requireAdmin(request, response)) return
    await handleListUsers(response)
    return
  }

  if (url.pathname === '/api/admin/users' && request.method === 'POST') {
    if (!await requireAdmin(request, response)) return
    await handleCreateUser(request, response)
    return
  }

  if (url.pathname.startsWith('/api/admin/users/') && request.method === 'PUT') {
    if (!await requireAdmin(request, response)) return
    const targetUserId = decodeURIComponent(url.pathname.slice('/api/admin/users/'.length))
    await handleUpdateUser(request, response, targetUserId, currentUser)
    return
  }

  if (url.pathname === '/api/settings/ai' && request.method === 'GET') {
    sendJson(response, 200, await loadAISettings())
    return
  }

  if (url.pathname === '/api/settings/ai/discover-models' && request.method === 'POST') {
    if (!await requireAdmin(request, response)) return
    const body = await readJsonBody(request)
    sendJson(response, 200, await fetchAvailableModels(body))
    return
  }

  if (url.pathname === '/api/settings/ai/models' && request.method === 'POST') {
    if (!await requireAdmin(request, response)) return
    const body = await readJsonBody(request)
    sendJson(response, 200, await addModelsToSettings(body))
    return
  }

  if (url.pathname === '/api/settings/ai/active-model' && request.method === 'PUT') {
    if (!await requireAdmin(request, response)) return
    const body = await readJsonBody(request)
    if (typeof body.modelId !== 'string' || !body.modelId.trim()) {
      sendError(response, 400, 'modelId is required')
      return
    }

    sendJson(response, 200, await setActiveModel(body.modelId.trim()))
    return
  }

  if (url.pathname.startsWith('/api/settings/ai/models/') && request.method === 'DELETE') {
    if (!await requireAdmin(request, response)) return
    const modelId = decodeURIComponent(url.pathname.slice('/api/settings/ai/models/'.length))
    sendJson(response, 200, await deleteModel(modelId))
    return
  }

  if (url.pathname === '/api/settings/ai' && request.method === 'DELETE') {
    if (!await requireAdmin(request, response)) return
    await clearAISettings()
    response.writeHead(204)
    response.end()
    return
  }

  if (url.pathname === '/api/llm/chat-completions' && request.method === 'POST') {
    await handleLLMChatCompletions(request, response)
    return
  }

  if (url.pathname === '/api/workbook/import-xlsx' && request.method === 'POST') {
    sendJson(response, 200, await importXlsxWorkbook(request))
    return
  }

  if (url.pathname === '/api/workbook/export-xlsx' && request.method === 'POST') {
    const body = await readJsonBody(request)
    const result = await exportXlsxWorkbook(body)
    sendBinary(response, 200, Buffer.from(result.buffer), {
      'Content-Type': mimeTypes['.xlsx'],
      'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(result.fileName)}`,
    })
    return
  }

  sendError(response, 404, 'API route not found')
}

async function serveStaticAsset(response, urlPathname) {
  const requestedPath = urlPathname === '/' ? 'index.html' : urlPathname.slice(1)
  const assetPath = path.resolve(distRoot, requestedPath)
  const relative = path.relative(distRoot, assetPath)
  if (relative.startsWith('..') || path.isAbsolute(relative)) {
    sendError(response, 403, 'Forbidden')
    return
  }

  try {
    const stats = await fs.stat(assetPath)
    if (stats.isDirectory()) {
      await serveStaticAsset(response, '/index.html')
      return
    }

    const ext = path.extname(assetPath)
    response.writeHead(200, { 'Content-Type': mimeTypes[ext] ?? 'application/octet-stream' })
    createReadStream(assetPath).pipe(response)
  } catch {
    const spaEntry = path.join(distRoot, 'index.html')
    try {
      await fs.access(spaEntry)
      response.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' })
      createReadStream(spaEntry).pipe(response)
    } catch {
      sendError(response, 404, 'Static asset not found')
    }
  }
}

const server = createServer(async (request, response) => {
  setCorsHeaders(request, response)

  if (!request.url) {
    sendError(response, 400, 'Missing request URL')
    return
  }

  const url = new URL(request.url, `http://${request.headers.host ?? `${host}:${port}`}`)

  try {
    if (url.pathname.startsWith('/api/')) {
      await handleApi(request, response, url)
      return
    }

    if (request.method === 'GET' || request.method === 'HEAD') {
      await serveStaticAsset(response, url.pathname)
      return
    }

    sendError(response, 404, 'Route not found')
  } catch (error) {
    console.error('[app-server]', error)
    sendError(response, 500, error instanceof Error ? error.message : 'Internal server error')
  }
})

await ensureDataDirectory()
await loadUsers()
await migrateLegacyAIConfigToAdmin()

server.listen(port, host, () => {
  console.log(`[app-server] listening at http://${host}:${port}`)
  console.log(`[app-server] ai config file: ${aiConfigPath}`)
})
