import type { FUniver } from '@univerjs/presets'

let _api: FUniver | null = null

export function setUniverAPI(api: FUniver): void {
  _api = api
}

export function getUniverAPI(): FUniver | null {
  return _api
}
