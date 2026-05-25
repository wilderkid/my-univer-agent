import { spawn } from 'node:child_process'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)
const projectRoot = path.resolve(__dirname, '..')
const serverEntry = path.join(projectRoot, 'server', 'index.js')

const env = { ...process.env }
const appProxy = env.APP_HTTP_PROXY?.trim()

if (appProxy) {
  env.HTTP_PROXY ??= appProxy
  env.HTTPS_PROXY ??= appProxy
}

env.NODE_USE_ENV_PROXY ??= '1'

const args = env.APP_SERVER_WATCH === '1'
  ? ['--watch', serverEntry]
  : [serverEntry]

const child = spawn(process.execPath, args, {
  cwd: projectRoot,
  env,
  stdio: 'inherit',
})

child.on('exit', (code, signal) => {
  if (signal) {
    process.kill(process.pid, signal)
    return
  }

  process.exit(code ?? 0)
})
