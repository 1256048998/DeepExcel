/**
 * 管理 API 客户端。
 *
 * 令牌存 sessionStorage 而非 localStorage：管理员令牌能禁用账号、开通付费套餐，
 * 关掉标签页就失效是合适的默认。服务端那边它也只有 8 小时寿命。
 */

const TOKEN_KEY = 'deepexcel.admin.token'

export function getToken(): string | null {
  return sessionStorage.getItem(TOKEN_KEY)
}

export function setToken(token: string | null): void {
  if (token) sessionStorage.setItem(TOKEN_KEY, token)
  else sessionStorage.removeItem(TOKEN_KEY)
}

export class ApiError extends Error {
  constructor(
    message: string,
    readonly status: number,
  ) {
    super(message)
  }
}

async function request<T>(path: string, init: RequestInit = {}): Promise<T> {
  const token = getToken()
  const response = await fetch(path, {
    ...init,
    headers: {
      ...(init.body ? { 'Content-Type': 'application/json' } : {}),
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
      ...(init.headers ?? {}),
    },
  })

  if (response.status === 401) {
    // 令牌过期或被吊销，回到登录页而不是让每个调用各自报错
    setToken(null)
    throw new ApiError('登录已失效，请重新登录', 401)
  }

  const text = await response.text()
  const body = text ? JSON.parse(text) : null

  if (!response.ok) {
    const detail = body?.detail
    const message =
      typeof detail === 'string' ? detail : detail?.message || `请求失败（${response.status}）`
    throw new ApiError(message, response.status)
  }
  return body as T
}

// ---------------------------------------------------------------------------

export interface Entitlement {
  plan: string
  status: string
  task_limit: number | null
  tasks_used: number
  tasks_remaining: number | null
  expires_at: string | null
}

export interface User {
  id: number
  email: string
  display_name: string | null
  status: string
  created_at: string
  entitlement: Entitlement | null
}

export interface InviteCode {
  id: number
  code: string
  note: string | null
  max_uses: number
  used_count: number
  expires_at: string | null
  disabled: boolean
  created_at: string
}

export interface Stats {
  total_users: number
  active_users: number
  installs_seen_7d: number
  tasks_7d: number
  task_success_rate_7d: number | null
  top_tool_errors_7d: { tool: string; count: number }[]
  client_versions: { version: string; count: number }[]
  // 自动更新健康度。上面的版本分布说的是客户端现在在哪个版本，说不了是不是
  // 更新机制把它们送过去的——一个悄悄失效的更新器看起来和"还没人升级"一样。
  update_upgrades_7d: number
  update_blocked_7d: number
  top_update_failures_7d: { reason: string; count: number }[]
}

export interface Order {
  order_no: string
  plan: string
  months: number
  amount_cents: number
  currency: string
  status: string
  channel: string | null
  created_at: string
  paid_at: string | null
}

export interface AuditEntry {
  id: number
  admin_id: number | null
  action: string
  target: string | null
  detail: unknown
  created_at: string
}

export const api = {
  login: (email: string, password: string) =>
    request<{ access_token: string; expires_at: number }>('/admin/api/auth/login', {
      method: 'POST',
      body: JSON.stringify({ email, password }),
    }),
  me: () => request<{ id: number; email: string }>('/admin/api/auth/me'),
  stats: () => request<Stats>('/admin/api/stats'),
  users: (query?: string) =>
    request<User[]>(`/admin/api/users${query ? `?q=${encodeURIComponent(query)}` : ''}`),
  setUserStatus: (id: number, status: 'active' | 'disabled') =>
    request<User>(`/admin/api/users/${id}/status`, {
      method: 'POST',
      body: JSON.stringify({ status }),
    }),
  updateEntitlement: (id: number, patch: Record<string, unknown>) =>
    request<User>(`/admin/api/users/${id}/entitlement`, {
      method: 'POST',
      body: JSON.stringify(patch),
    }),
  invites: () => request<InviteCode[]>('/admin/api/invites'),
  createInvite: (note: string, maxUses: number) =>
    request<InviteCode>('/admin/api/invites', {
      method: 'POST',
      body: JSON.stringify({ note: note || null, max_uses: maxUses }),
    }),
  disableInvite: (id: number) =>
    request<InviteCode>(`/admin/api/invites/${id}/disable`, { method: 'POST' }),
  orders: () => request<Order[]>('/admin/api/orders'),
  markPaid: (orderNo: string) =>
    request<Order>(`/admin/api/orders/${orderNo}/mark-paid`, { method: 'POST' }),
  audit: () => request<AuditEntry[]>('/admin/api/audit'),
}
