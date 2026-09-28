// 错误卡片下面给哪些按钮：按侧车分类的 error.code（docs/ui-event-protocol.md「error.code」）。
// 只是界面上的下一步入口；走托管还是自带 Key 仍由服务端决定，这里不判断。
// 纯函数，有测试：errorActions.test.ts

export type ErrorAction = 'retry' | 'settings' | 'account' | 'new_chat' | 'copy_detail'

export const ERROR_ACTION_TEXT: Record<ErrorAction, string> = {
  retry: '重试',
  settings: '打开模型设置',
  account: '查看账号额度',
  new_chat: '开新对话',
  copy_detail: '复制诊断信息',
}

type ErrorLike = { code: string; retryable?: boolean; detail?: string }

export function errorActions(error: ErrorLike, canRetry: boolean): ErrorAction[] {
  const actions: ErrorAction[] = []
  if (error.retryable && canRetry) actions.push('retry')
  if (error.code === 'auth' || error.code === 'model_not_found') actions.push('settings')
  if (error.code === 'quota' || error.code === 'task_limit') actions.push('account')
  if (error.code === 'context_too_long') actions.push('new_chat')
  if (error.detail) actions.push('copy_detail')
  return actions
}

/** 重试要重发的那句话：错误卡片之前最近的一条用户消息。 */
export function retryTarget(messages: { role: string; content: string }[], errorIndex: number): string | null {
  for (let i = Math.min(errorIndex, messages.length) - 1; i >= 0; i--) {
    if (messages[i].role === 'user' && messages[i].content.trim()) return messages[i].content
  }
  return null
}
