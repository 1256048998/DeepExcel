// 复制文字：优先 Clipboard API；面板没拿到焦点（刚从 Excel 单元格点过来）时它会拒绝，
// 退回隐藏 textarea + execCommand。返回是否复制成功，调用方据此给「已复制」反馈。
export async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text)
    return true
  } catch {
    try {
      const ta = document.createElement('textarea')
      ta.value = text
      ta.setAttribute('readonly', '')
      ta.style.position = 'fixed'
      ta.style.opacity = '0'
      document.body.appendChild(ta)
      ta.select()
      const ok = document.execCommand('copy')
      ta.remove()
      return ok
    } catch {
      return false
    }
  }
}
