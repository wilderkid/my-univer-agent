import { UniverAgent } from './univer-agent'

export type LLMProvider = 'qwen' | 'openai' | 'custom'

export interface AgentConfig {
  apiKey: string
  provider?: LLMProvider
  model?: string
  baseURL?: string
}

const PROVIDER_DEFAULTS: Record<Exclude<LLMProvider, 'custom'>, { model: string; baseURL: string }> = {
  qwen: {
    model: 'qwen-plus',
    baseURL: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
  },
  openai: {
    model: 'gpt-4o-mini',
    baseURL: 'https://api.openai.com/v1',
  },
}

export function createAgent(config: AgentConfig): UniverAgent {
  const { apiKey, provider = 'qwen', model, baseURL } = config

  let resolvedModel: string
  let resolvedBaseURL: string

  if (provider === 'custom') {
    resolvedModel = model ?? ''
    resolvedBaseURL = baseURL ?? ''
  } else {
    const defaults = PROVIDER_DEFAULTS[provider]
    resolvedModel = model?.trim() || defaults.model
    resolvedBaseURL = baseURL?.trim() || defaults.baseURL
  }

  return new UniverAgent({ apiKey, model: resolvedModel, baseURL: resolvedBaseURL })
}

let _agent: UniverAgent | null = null

export function getAgent(): UniverAgent | null {
  return _agent
}

export function initAgent(config: AgentConfig): UniverAgent {
  _agent = createAgent(config)
  return _agent
}

export function destroyAgent(): void {
  _agent = null
}
