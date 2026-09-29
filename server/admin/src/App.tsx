import { Fragment, useCallback, useEffect, useState } from 'react'
import { api, ApiError, getToken, setToken } from './api'
import type { AuditEntry, Entitlement, InviteCode, Order, Stats, User } from './api'

type Tab = 'dashboard' | 'users' | 'invites' | 'orders' | 'audit'

const PLAN_LABELS: Record<string, string> = {
  beta: '内测',
  free: '免费',
  pro: '专业版',
  team: '团队版',
  byok: '自带密钥',
}

function formatTime(value: string | null): string {
  if (!value) return '—'
  return new Date(value).toLocaleString('zh-CN', { hour12: false })
}

// 与 server/app/routers/billing.py 的 CATALOG 一致：这两档付费后走托管转发。
// 这里只用来提前说明，真正拦截的是服务端。
const HOSTED_PLANS = new Set(['pro', 'team'])

// 托管转发是否已开通。null 表示还没问到：此时什么都不禁用，服务端的 409 兜底。
function useHostedAvailable(): boolean | null {
  const [available, setAvailable] = useState<boolean | null>(null)
  useEffect(() => {
    api
      .meta()
      .then((meta) => setAvailable(meta.hosted_routing_available))
      .catch(() => setAvailable(null))
  }, [])
  return available
}

function formatMoney(cents: number, currency: string): string {
  return `${currency === 'CNY' ? '¥' : '$'}${(cents / 100).toFixed(2)}`
}

// ---------------------------------------------------------------------------

function Login({ onDone }: { onDone: () => void }) {
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')

  const submit = async () => {
    setBusy(true)
    setError('')
    try {
      const result = await api.login(email.trim(), password)
      setToken(result.access_token)
      onDone()
    } catch (e) {
      setError(e instanceof Error ? e.message : '登录失败')
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="login">
      <h1>DeepExcel 运营后台</h1>
      <input
        placeholder="管理员邮箱"
        value={email}
        onChange={(e) => setEmail(e.target.value)}
        disabled={busy}
      />
      <input
        type="password"
        placeholder="密码"
        value={password}
        onChange={(e) => setPassword(e.target.value)}
        onKeyDown={(e) => e.key === 'Enter' && !busy && void submit()}
        disabled={busy}
      />
      {error && <p className="error">{error}</p>}
      <button onClick={() => void submit()} disabled={busy}>
        {busy ? '登录中…' : '登录'}
      </button>
      <p className="hint">
        首个管理员由环境变量 BOOTSTRAP_ADMIN_EMAIL / BOOTSTRAP_ADMIN_PASSWORD 创建。
        没有公开的注册入口。
      </p>
    </div>
  )
}

// ---------------------------------------------------------------------------

function Dashboard() {
  const [stats, setStats] = useState<Stats | null>(null)
  const [error, setError] = useState('')

  useEffect(() => {
    api.stats().then(setStats).catch((e) => setError(e.message))
  }, [])

  if (error) return <p className="error">{error}</p>
  if (!stats) return <p className="muted">加载中…</p>

  return (
    <div>
      <div className="cards">
        <Card label="注册用户" value={String(stats.total_users)} />
        <Card label="活跃账号" value={String(stats.active_users)} />
        <Card label="7 天内活跃安装" value={String(stats.installs_seen_7d)} />
        <Card label="7 天任务数" value={String(stats.tasks_7d)} />
        <Card
          label="任务一次成功率"
          value={
            stats.task_success_rate_7d === null
              ? '暂无数据'
              : `${(stats.task_success_rate_7d * 100).toFixed(1)}%`
          }
          // 北极星指标：所有能力改动都该对着它评判
          highlight
        />
        <Card label="7 天成功升级" value={String(stats.update_upgrades_7d)} />
        <Card
          label="更新装不上（已放弃）"
          value={String(stats.update_blocked_7d)}
          // 0 是正常，非 0 需要有人去处理：这些用户的客户端已经下载并验签
          // 成功、装了三次都失败、不再提示了，多半被安全软件拦掉。看板上没有
          // 第二个指标会反映这件事，所以它必须自己喊出来。
          warn={stats.update_blocked_7d > 0}
        />
      </div>

      <h3>工具失败 TOP 10（7 天）</h3>
      {stats.top_tool_errors_7d.length === 0 ? (
        <p className="muted">暂无失败记录</p>
      ) : (
        <table>
          <thead>
            <tr>
              <th>工具</th>
              <th>失败次数</th>
            </tr>
          </thead>
          <tbody>
            {stats.top_tool_errors_7d.map((row) => (
              <tr key={row.tool}>
                <td className="mono">{row.tool}</td>
                <td>{row.count}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}

      <h3>更新失败原因 TOP 10（7 天）</h3>
      {stats.top_update_failures_7d.length === 0 ? (
        <p className="muted">暂无失败记录</p>
      ) : (
        <table>
          <thead>
            <tr>
              <th>分类码</th>
              <th>次数</th>
            </tr>
          </thead>
          <tbody>
            {stats.top_update_failures_7d.map((row) => (
              <tr key={row.reason}>
                <td className="mono">{row.reason}</td>
                <td>{row.count}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}

      <h3>客户端版本分布（7 天）</h3>
      {stats.client_versions.length === 0 ? (
        <p className="muted">暂无数据</p>
      ) : (
        <table>
          <thead>
            <tr>
              <th>版本</th>
              <th>安装数</th>
            </tr>
          </thead>
          <tbody>
            {stats.client_versions.map((row) => (
              <tr key={row.version}>
                <td className="mono">{row.version}</td>
                <td>{row.count}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  )
}

function Card({ label, value, highlight, warn }: {
  label: string; value: string; highlight?: boolean; warn?: boolean
}) {
  // warn 和 highlight 是两件事：highlight 表示"这个数最重要"，
  // warn 表示"这个数现在不对，需要人去做点什么"。
  const className = warn ? 'card card-warn' : highlight ? 'card card-highlight' : 'card'
  return (
    <div className={className}>
      <span className="card-label">{label}</span>
      <span className="card-value">{value}</span>
    </div>
  )
}

// ---------------------------------------------------------------------------

const PAGE_SIZE = 50

const ENTITLEMENT_STATUS_LABELS: Record<string, string> = {
  active: '有效',
  expired: '已过期',
  suspended: '已暂停',
}

function formatDate(value: string): string {
  return new Date(value).toLocaleDateString('zh-CN')
}

// <input type="date"> 用本地日期，和上面的显示保持一致
function toDateInput(value: string | null): string {
  if (!value) return ''
  const date = new Date(value)
  const pad = (n: number) => String(n).padStart(2, '0')
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`
}

// 到期日当天（本地时间）结束前都算有效
function endOfDay(date: string): string {
  return new Date(`${date}T23:59:59`).toISOString()
}

function EntitlementEditor({ entitlement, onSave, onCancel }: {
  entitlement: Entitlement
  onSave: (patch: Record<string, unknown>) => Promise<boolean>
  onCancel: () => void
}) {
  const [plan, setPlan] = useState(entitlement.plan)
  const [status, setStatus] = useState(entitlement.status)
  const [limit, setLimit] = useState(entitlement.task_limit === null ? '' : String(entitlement.task_limit))
  const [expiry, setExpiry] = useState(toDateInput(entitlement.expires_at))
  const [busy, setBusy] = useState(false)
  const [problem, setProblem] = useState('')

  const save = async () => {
    const newLimit = limit.trim() === '' ? null : Number(limit)
    if (newLimit !== null && (!Number.isInteger(newLimit) || newLimit < 0)) {
      setProblem('额度填 0 或正整数；留空表示不限')
      return
    }
    // 只发改了的字段：服务端只改收到的字段，null 对额度和到期表示"不限"
    const patch: Record<string, unknown> = {}
    if (plan !== entitlement.plan) patch.plan = plan
    if (status !== entitlement.status) patch.status = status
    if (newLimit !== entitlement.task_limit) patch.task_limit = newLimit
    if (expiry !== toDateInput(entitlement.expires_at)) patch.expires_at = expiry ? endOfDay(expiry) : null
    if (Object.keys(patch).length === 0) {
      onCancel()
      return
    }
    setBusy(true)
    setProblem('')
    await onSave(patch)
    setBusy(false)
  }

  return (
    <div className="editor">
      <label>
        套餐
        <select value={plan} onChange={(e) => setPlan(e.target.value)} disabled={busy}>
          {Object.entries(PLAN_LABELS).map(([key, label]) => (
            <option key={key} value={key}>{label}</option>
          ))}
        </select>
      </label>
      <label>
        权益状态
        <select value={status} onChange={(e) => setStatus(e.target.value)} disabled={busy}>
          {Object.entries(ENTITLEMENT_STATUS_LABELS).map(([key, label]) => (
            <option key={key} value={key}>{label}</option>
          ))}
        </select>
      </label>
      <label>
        每月额度（点）
        <input
          type="number"
          min={0}
          placeholder="不限"
          value={limit}
          onChange={(e) => setLimit(e.target.value)}
          disabled={busy}
        />
      </label>
      <label>
        到期日
        <input type="date" value={expiry} onChange={(e) => setExpiry(e.target.value)} disabled={busy} />
      </label>
      <div className="editor-actions">
        <button onClick={() => void save()} disabled={busy}>{busy ? '保存中…' : '保存'}</button>
        <button onClick={onCancel} disabled={busy}>取消</button>
      </div>
      {problem && <p className="error">{problem}</p>}
      <p className="hint">
        额度留空 = 不限，到期日留空 = 不过期。这里改套餐不收费、也不改出口，适合手工赠送或补偿；
        用户付费请走订单。额度只对托管转发的用户生效，已用点数不会被清零。每个任务按所用模型扣点（DeepSeek 1 点）。每次保存都记入审计。
      </p>
    </div>
  )
}

function Users() {
  const [users, setUsers] = useState<User[]>([])
  const [hasMore, setHasMore] = useState(false)
  const [page, setPage] = useState(0)
  const [query, setQuery] = useState('')
  const [search, setSearch] = useState('')
  const [editing, setEditing] = useState<number | null>(null)
  const [error, setError] = useState('')
  const hostedAvailable = useHostedAvailable()

  const load = useCallback(async () => {
    try {
      // 多取一条来判断有没有下一页：接口不返回总数
      const rows = await api.users(search, page * PAGE_SIZE, PAGE_SIZE + 1)
      setUsers(rows.slice(0, PAGE_SIZE))
      setHasMore(rows.length > PAGE_SIZE)
      setError('')
    } catch (e) {
      setError(e instanceof Error ? e.message : '加载失败')
    }
  }, [search, page])

  useEffect(() => {
    void load()
  }, [load])

  const runSearch = () => {
    setEditing(null)
    const next = query.trim()
    if (page === 0 && next === search) void load()
    else {
      setPage(0)
      setSearch(next)
    }
  }

  const goTo = (target: number) => {
    setEditing(null)
    setPage(target)
  }

  const act = async (work: Promise<unknown>): Promise<boolean> => {
    try {
      await work
      await load()
      return true
    } catch (e) {
      setError(e instanceof Error ? e.message : '操作失败')
      return false
    }
  }

  const toggleStatus = (user: User) => {
    if (
      user.status === 'active' &&
      !window.confirm(`停用 ${user.email}？\n\n该用户会立即退出登录，恢复之前无法使用。`)
    ) {
      return
    }
    void act(api.setUserStatus(user.id, user.status === 'active' ? 'disabled' : 'active'))
  }

  return (
    <div>
      <div className="toolbar">
        <input
          placeholder="按邮箱搜索"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && runSearch()}
        />
        <button onClick={runSearch}>搜索</button>
      </div>
      {error && <p className="error">{error}</p>}

      <table>
        <thead>
          <tr>
            <th>邮箱</th>
            <th>状态</th>
            <th>套餐</th>
            <th>出口</th>
            <th>额度</th>
            <th>注册时间</th>
            <th>操作</th>
          </tr>
        </thead>
        <tbody>
          {users.map((user) => {
            const entitlement = user.entitlement
            const expired =
              entitlement?.expires_at != null && new Date(entitlement.expires_at) <= new Date()
            return (
              <Fragment key={user.id}>
                <tr>
                  <td>{user.email}</td>
                  <td className={user.status === 'active' ? 'ok' : 'warn'}>
                    {user.status === 'active' ? '正常' : '已停用'}
                  </td>
                  <td>
                    {PLAN_LABELS[entitlement?.plan ?? ''] ?? entitlement?.plan ?? '—'}
                    {entitlement && entitlement.status !== 'active' && (
                      <span className="warn">
                        {' · '}
                        {ENTITLEMENT_STATUS_LABELS[entitlement.status] ?? entitlement.status}
                      </span>
                    )}
                    {entitlement?.expires_at && (
                      <div className={expired ? 'warn' : 'muted'}>
                        {expired ? '已于 ' : '到期 '}
                        {formatDate(entitlement.expires_at)}
                        {expired ? ' 到期' : ''}
                      </div>
                    )}
                  </td>
                  <td>
                    <select
                      value={entitlement?.routing_mode ?? 'byok'}
                      onChange={(e) =>
                        void act(api.updateEntitlement(user.id, { routing_mode: e.target.value }))
                      }
                    >
                      <option value="byok">自带密钥</option>
                      {/* 托管没开时切过去，这个用户下次刷新出口配置就会被拒、完全用不了 */}
                      <option value="hosted" disabled={hostedAvailable === false}>
                        {hostedAvailable === false ? '托管转发（未开通）' : '托管转发'}
                      </option>
                    </select>
                    {hostedAvailable === false && entitlement?.routing_mode === 'hosted' && (
                      <div className="warn">托管未开通，该用户现在无法使用，请切回自带密钥</div>
                    )}
                  </td>
                  <td>
                    {entitlement?.task_limit === null || entitlement == null
                      ? '不限'
                      : `${entitlement.tasks_used} / ${entitlement.task_limit} 点`}
                  </td>
                  <td className="muted">{formatTime(user.created_at)}</td>
                  <td className="actions">
                    {entitlement && (
                      <button onClick={() => setEditing(editing === user.id ? null : user.id)}>
                        {editing === user.id ? '收起' : '编辑'}
                      </button>
                    )}
                    <button onClick={() => toggleStatus(user)}>
                      {user.status === 'active' ? '停用' : '恢复'}
                    </button>
                  </td>
                </tr>
                {editing === user.id && entitlement && (
                  <tr>
                    <td colSpan={7}>
                      <EntitlementEditor
                        entitlement={entitlement}
                        onCancel={() => setEditing(null)}
                        onSave={async (patch) => {
                          const ok = await act(api.updateEntitlement(user.id, patch))
                          if (ok) setEditing(null)
                          return ok
                        }}
                      />
                    </td>
                  </tr>
                )}
              </Fragment>
            )
          })}
        </tbody>
      </table>
      {users.length === 0 && <p className="muted">没有用户</p>}
      {(page > 0 || hasMore) && (
        <div className="pager">
          <button disabled={page === 0} onClick={() => goTo(page - 1)}>上一页</button>
          <span className="muted">第 {page + 1} 页</span>
          <button disabled={!hasMore} onClick={() => goTo(page + 1)}>下一页</button>
        </div>
      )}
    </div>
  )
}

// ---------------------------------------------------------------------------

function Invites() {
  const [invites, setInvites] = useState<InviteCode[]>([])
  const [note, setNote] = useState('')
  const [maxUses, setMaxUses] = useState(1)
  const [error, setError] = useState('')

  const load = useCallback(async () => {
    try {
      setInvites(await api.invites())
      setError('')
    } catch (e) {
      setError(e instanceof Error ? e.message : '加载失败')
    }
  }, [])

  useEffect(() => {
    void load()
  }, [load])

  return (
    <div>
      <div className="toolbar">
        <input placeholder="备注（给谁用）" value={note} onChange={(e) => setNote(e.target.value)} />
        <input
          type="number"
          min={1}
          max={1000}
          value={maxUses}
          onChange={(e) => setMaxUses(Number(e.target.value))}
        />
        <button
          onClick={async () => {
            try {
              await api.createInvite(note, maxUses)
              setNote('')
              await load()
            } catch (e) {
              setError(e instanceof Error ? e.message : '创建失败')
            }
          }}
        >
          生成邀请码
        </button>
      </div>
      {error && <p className="error">{error}</p>}

      <table>
        <thead>
          <tr>
            <th>邀请码</th>
            <th>备注</th>
            <th>已用 / 上限</th>
            <th>状态</th>
            <th>创建时间</th>
            <th />
          </tr>
        </thead>
        <tbody>
          {invites.map((invite) => (
            <tr key={invite.id}>
              <td className="mono">{invite.code}</td>
              <td>{invite.note || '—'}</td>
              <td>
                {invite.used_count} / {invite.max_uses}
              </td>
              <td className={invite.disabled ? 'warn' : 'ok'}>
                {invite.disabled
                  ? '已停用'
                  : invite.used_count >= invite.max_uses
                    ? '已用完'
                    : '可用'}
              </td>
              <td className="muted">{formatTime(invite.created_at)}</td>
              <td>
                {!invite.disabled && (
                  <button
                    onClick={async () => {
                      if (
                        !window.confirm(
                          `停用邀请码 ${invite.code}？\n\n已经用它注册的用户不受影响，但这个码不能再用来注册。`,
                        )
                      ) {
                        return
                      }
                      try {
                        await api.disableInvite(invite.id)
                        await load()
                      } catch (e) {
                        setError(e instanceof Error ? e.message : '操作失败')
                      }
                    }}
                  >
                    停用
                  </button>
                )}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      {invites.length === 0 && <p className="muted">还没有邀请码</p>}
    </div>
  )
}

// ---------------------------------------------------------------------------

function Orders() {
  const [orders, setOrders] = useState<Order[]>([])
  const [error, setError] = useState('')
  const hostedAvailable = useHostedAvailable()

  const load = useCallback(async () => {
    try {
      setOrders(await api.orders())
      setError('')
    } catch (e) {
      setError(e instanceof Error ? e.message : '加载失败')
    }
  }, [])

  useEffect(() => {
    void load()
  }, [load])

  return (
    <div>
      <p className="hint">
        支付渠道尚未接入（微信支付 / 支付宝都需要企业商户号）。收到款后在这里手工标记为已支付，
        权益会立即生效。每次标记都会写入审计日志。
      </p>
      {hostedAvailable === false && (
        <p className="hint warn">
          托管转发尚未开通：专业版 / 团队版付费后走托管，现在标记已支付会让用户立即无法使用，
          所以暂不能标记，服务端也不再接受这两档的新订单。「自带密钥」不受影响。
        </p>
      )}
      {error && <p className="error">{error}</p>}
      <table>
        <thead>
          <tr>
            <th>订单号</th>
            <th>套餐</th>
            <th>时长</th>
            <th>金额</th>
            <th>状态</th>
            <th>创建时间</th>
            <th />
          </tr>
        </thead>
        <tbody>
          {orders.map((order) => (
            <tr key={order.order_no}>
              <td className="mono">{order.order_no}</td>
              <td>{PLAN_LABELS[order.plan] ?? order.plan}</td>
              <td>{order.months} 个月</td>
              <td>{formatMoney(order.amount_cents, order.currency)}</td>
              <td className={order.status === 'paid' ? 'ok' : 'muted'}>
                {order.status === 'paid' ? '已支付' : order.status === 'pending' ? '待支付' : order.status}
              </td>
              <td className="muted">{formatTime(order.created_at)}</td>
              <td>
                {order.status === 'pending' &&
                  (hostedAvailable === false && HOSTED_PLANS.has(order.plan) ? (
                    <span className="muted">托管未开通</span>
                  ) : (
                    <button
                      onClick={async () => {
                        const plan = PLAN_LABELS[order.plan] ?? order.plan
                        if (
                          !window.confirm(
                            `确认已收到 ${formatMoney(order.amount_cents, order.currency)}？\n\n` +
                              `订单 ${order.order_no}：${plan} ${order.months} 个月。标记后权益立即生效，` +
                              '记入审计日志，无法撤销。',
                          )
                        ) {
                          return
                        }
                        try {
                          await api.markPaid(order.order_no)
                          await load()
                        } catch (e) {
                          setError(e instanceof Error ? e.message : '操作失败')
                        }
                      }}
                    >
                      标记已支付
                    </button>
                  ))}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      {orders.length === 0 && <p className="muted">还没有订单</p>}
    </div>
  )
}

// ---------------------------------------------------------------------------

function Audit() {
  const [entries, setEntries] = useState<AuditEntry[]>([])
  const [error, setError] = useState('')

  useEffect(() => {
    // 加载失败要说出来：显示成"暂无记录"会让人以为什么都没发生过
    api
      .audit()
      .then(setEntries)
      .catch((e) => setError(e instanceof Error ? e.message : '加载失败'))
  }, [])

  return (
    <div>
      {error && <p className="error">{error}</p>}
      <p className="hint">所有管理操作都会记录在这里——"是谁禁用了这个账号"需要有答案。</p>
      <table>
        <thead>
          <tr>
            <th>时间</th>
            <th>操作</th>
            <th>对象</th>
            <th>详情</th>
          </tr>
        </thead>
        <tbody>
          {entries.map((entry) => (
            <tr key={entry.id}>
              <td className="muted">{formatTime(entry.created_at)}</td>
              <td className="mono">{entry.action}</td>
              <td>{entry.target || '—'}</td>
              <td className="mono muted">
                {entry.detail ? JSON.stringify(entry.detail) : '—'}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      {entries.length === 0 && !error && <p className="muted">暂无记录</p>}
    </div>
  )
}

// ---------------------------------------------------------------------------

export function App() {
  const [authed, setAuthed] = useState(Boolean(getToken()))
  const [tab, setTab] = useState<Tab>('dashboard')
  const [email, setEmail] = useState('')

  useEffect(() => {
    if (!authed) return
    api
      .me()
      .then((me) => setEmail(me.email))
      .catch((e) => {
        if (e instanceof ApiError && e.status === 401) setAuthed(false)
      })
  }, [authed])

  if (!authed) return <Login onDone={() => setAuthed(true)} />

  const tabs: [Tab, string][] = [
    ['dashboard', '概览'],
    ['users', '用户'],
    ['invites', '邀请码'],
    ['orders', '订单'],
    ['audit', '审计'],
  ]

  return (
    <div className="app">
      <header>
        <h1>DeepExcel 运营后台</h1>
        <div className="header-right">
          <span className="muted">{email}</span>
          <button
            onClick={() => {
              setToken(null)
              setAuthed(false)
            }}
          >
            退出
          </button>
        </div>
      </header>

      <nav>
        {tabs.map(([key, label]) => (
          <button key={key} className={tab === key ? 'active' : ''} onClick={() => setTab(key)}>
            {label}
          </button>
        ))}
      </nav>

      <main>
        {tab === 'dashboard' && <Dashboard />}
        {tab === 'users' && <Users />}
        {tab === 'invites' && <Invites />}
        {tab === 'orders' && <Orders />}
        {tab === 'audit' && <Audit />}
      </main>
    </div>
  )
}
