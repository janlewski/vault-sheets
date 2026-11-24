// Modified 2026-11-24 by janlewski for Vault Sheets.

export function pickFile(accept: string): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = accept
    let settled = false
    const finish = (file: File | null) => {
      if (settled)
        return
      settled = true
      input.remove()
      resolve(file)
    }
    input.addEventListener('change', () => {
      finish(input.files?.[0] ?? null)
    })
    input.addEventListener('cancel', () => {
      finish(null)
    })
    document.body.appendChild(input)
    input.click()
  })
}

export function downloadBlob(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = filename
  document.body.appendChild(anchor)
  anchor.click()
  anchor.remove()
  URL.revokeObjectURL(url)
}

export function downloadStem(workbookName: string): string {
  const base = workbookName.split(/[/\\]/).pop() || 'sheet'
  const stem = base.replace(/\.(sheet|xlsx|csv|md)$/i, '')
  return stem || 'sheet'
}
