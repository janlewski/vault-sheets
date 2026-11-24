// Modified 2026-11-24 by janlewski for Vault Sheets.

import { useEffect, useRef } from 'react'
import { onEvent } from './eventBus'

export function useEventBus<T>(event: string, handler: (payload: T) => void): void {
  const handlerRef = useRef(handler)
  handlerRef.current = handler

  useEffect(() => {
    return onEvent(event, payload => handlerRef.current(payload as T))
  }, [event])
}
