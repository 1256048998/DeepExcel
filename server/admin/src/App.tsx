import { useCallback, useEffect, useState } from 'react'
import { api, ApiError, getToken, setToken } from './api'
import type { AuditEntry, InviteCode, Order, Stats, User } from './api'

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

function Users() {
  const [users, setUsers] = useState<User[]>([])
  const [query, setQuery] = useState('')
  const [error, setError] = useState('')

  const load = useCallback(async (q?: string) => {
    try {
      setUsers(await api.users(q))
      setError('')
    } catch (e) {
      setError(e instanceof Error ? e.message : '加载失败')
    }
  }, [])

  useEffect(() => {
    void load()
  }, [load])

  const act = async (work: Promise<unknown>) => {
    try {
      await work
      await load(query)
    } catch (e) {
      setError(e instanceof Error ? e.message : '操作失败')
    }
  }

  return (
    <div>
      <div className="toolbar">
        <input
          placeholder="按邮箱搜索"
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && void load(query)}
        />
        <button onClick={() => void load(query)}>搜索</button>
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
          {users.map((user) => (
            <tr key={user.id}>
              <td>{user.email}</td>
              <td className={user.status === 'active' ? 'ok' : 'warn'}>
                {user.status === 'active' ? '正常' : '已停用'}
              </td>
              <td>{PLAN_LABELS[user.entitlement?.plan ?? ''] ?? user.entitlement?.plan ?? '—'}</td>
              <td>
                <select
                  value={(user.entitlement as { routing_mode?: string })?.routing_mode ?? 'byok'}
                  onChange={(e) =>
                    void act(api.updateEntitlement(user.id, { routing_mode: e.target.value }))
                  }
                >
                  <option value="byok">自带密钥</option>
                  <option value="hosted">托管转发</option>
                </select>
              </td>
              <td>
                {user.entitlement?.task_limit === null || user.entitlement == null
                  ? '不限'
                  : `${user.entitlement.tasks_used} / ${user.entitlement.task_limit}`}
              </td>
              <td className="muted">{formatTime(user.created_at)}</td>
              <td>
                <button
                  onClick={() =>
                    void act(
                      api.setUserStatus(user.id, user.status === 'active' ? 'disabled' : 'active'),
                    )
                  }
                >
                  {user.status === 'active' ? '停用' : '恢复'}
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      {users.length === 0 && <p className="muted">没有用户</p>}
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
                      await api.disableInvite(invite.id)
                      await load()
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
                {order.status === 'pending' && (
                  <button
                    onClick={async () => {
                      await api.markPaid(order.order_no)
                      await load()
                    }}
                  >
                    标记已支付
                  </button>
                )}
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

  useEffect(() => {
    api.audit().then(setEntries).catch(() => setEntries([]))
  }, [])

  return (
    <div>
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
      {entries.length === 0 && <p className="muted">暂无记录</p>}
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
