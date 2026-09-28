// 面板里的全局快捷键（Esc 停止、确认抽屉的 Enter / Esc）要避开两种情况：
//   1. 焦点在输入框里：任务进行中也能打字插话，那时的 Enter 是发插话，不是「允许」
//   2. 有弹窗 / 下拉 / 菜单开着：Esc 先用来关它们
// 纯函数，有测试：keys.test.ts

export function isEditable(target: EventTarget | null): boolean {
  const el = target as HTMLElement | null
  if (!el || typeof el.tagName !== 'string') return false
  const tag = el.tagName.toLowerCase()
  return tag === 'textarea' || (tag === 'input' && !['checkbox', 'radio', 'button'].includes((el as HTMLInputElement).type))
    || el.isContentEditable === true
}

// 这些元素在的时候，Esc 属于它们
export const ESC_OWNERS = '.config-overlay, .conv-panel-overlay, .history-panel-overlay, .import-overlay, '
  + '.welcome-overlay, .header-menu-list, .model-picker-pop, .prompt-dropdown'

export function escOwnedByOverlay(root: ParentNode = document): boolean {
  return root.querySelector(ESC_OWNERS) !== null
}
