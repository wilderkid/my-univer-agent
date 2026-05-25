import { spawn } from 'node:child_process'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)
const projectRoot = path.resolve(__dirname, '..')

const children = []
let shuttingDown = false

function startProcess(commandLine, extraEnv = {}) {
  const child = spawn(commandLine, {
    cwd: projectRoot,
    env: { ...process.env, ...extraEnv },
    shell: true,
    stdio: 'inherit',
  })

  children.push(child)

  child.on('exit', (code) => {
    if (shuttingDown) return
    shuttingDown = true
    stopChildren()
    process.exit(code ?? 0)
  })

  return child
}

function stopChildren() {
  for (const child of children) {
    if (!child.killed) {
      child.kill('SIGTERM')
    }
  }
}

process.on('SIGINT', () => {
  shuttingDown = true
  stopChildren()
  process.exit(0)
})

process.on('SIGTERM', () => {
  shuttingDown = true
  stopChildren()
  process.exit(0)
})

startProcess('npm run dev:server', { APP_SERVER_WATCH: '1' })
startProcess('npm run dev:vite')
