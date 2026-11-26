// Modified 2026-11-24 by janlewski for Vault Sheets.

type EventHandler = (payload: unknown) => void

const listeners = new Map<string, Set<EventHandler>>()

export function emitEvent(event: string, payload?: unknown): void {
  const handlers = listeners.get(event)
  if (!handlers)
    return
  for (const handler of [...handlers])
    handler(payload)
}

export function onEvent(event: string, handler: EventHandler): () => void {
  let handlers = listeners.get(event)
  if (!handlers) {
    handlers = new Set()
    listeners.set(event, handlers)
  }
  handlers.add(handler)
  return () => {
    handlers.delete(handler)
    if (handlers.size === 0)
      listeners.delete(event)
  }
}
