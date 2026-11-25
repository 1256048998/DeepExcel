import type { HostMessage } from '../bridge'
import type { AccountStatus } from '../components/AccountPanel'

/**
 * 宿主对 account_* 请求的回复怎么读。
 *
 * 宿主失败时回的是 `{ type: 'error', payload: { message } }`（邮箱或密码不正确、需要邀请码、
 * 连不上服务器……）。面板之前把任何回复都当成 AccountStatus 用：错误对象被当成了
 * 登录状态，state 是 undefined，于是表单原样留着、也没有错误提示——用户看到的就是
 * 「点了没反应」。这里把三种情况分开：拿到了状态、拿到了错误、根本没回。
 */
export type AccountReply =
  | { ok: true; status: AccountStatus }
  | { ok: false; error: string }

const STATES: ReadonlyArray<AccountStatus['state']> = ['signedout', 'signedin', 'offline', 'expired']

/** 宿主没起会话时曾回过 'signed_out'，和面板认的 'signedout' 对不上；两种都认。 */
export function normalizeAccountState(state: unknown): AccountStatus['state'] | null {
  if (state === 'signed_out') return 'signedout'
  return typeof state === 'string' && (STATES as readonly string[]).includes(state)
    ? (state as AccountStatus['state'])
    : null
}

export function readAccountReply(resp: HostMessage | null | undefined, fallbackError: string): AccountReply {
  if (!resp) return { ok: false, error: '服务器没有响应，请检查网络后重试' }
  if (resp.type === 'error') {
    const message = typeof resp.payload?.message === 'string' ? resp.payload.message.trim() : ''
    return { ok: false, error: message || fallbackError }
  }
  if (resp.type === 'account_status' && resp.payload) {
    const state = normalizeAccountState(resp.payload.state)
    if (state) return { ok: true, status: { ...resp.payload, state } as AccountStatus }
  }
  return { ok: false, error: fallbackError }
}

/**
 * 是否要邀请码。只有服务器明确说「不需要」才隐藏邀请码输入框；
 * 探测失败（连不上、超时）时返回 null，让输入框以「如服务器要求」的样子留着——
 * 之前探测失败被读成了「不需要」，输入框消失，注册必然被服务器以缺邀请码拒绝。
 */
export function readInviteRequired(resp: HostMessage | null | undefined): boolean | null {
  if (resp?.type !== 'account_server_meta' || !resp.payload) return null
  const value = resp.payload.invite_required
  return typeof value === 'boolean' ? value : null
}
