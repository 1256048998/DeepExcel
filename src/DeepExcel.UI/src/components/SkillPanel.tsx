import { useState, useEffect, useCallback } from 'react'
import { sendToHostWithResponse } from '../bridge'

/**
 * 技能库。
 *
 * 技能是「上次成功做成的一串操作」，不是宏。重放仍然走 Agent，也仍然会经过
 * 同样的确认与变更预览——如果重放能跳过这些检查，技能库就变成了绕过它们的
 * 途径。
 */

export interface SkillParameterView {
  name: string
  kind: 'range' | 'sheet' | 'date' | 'file' | 'number' | 'text'
  label: string
  default_value: string
}

export interface SkillView {
  id: string
  name: string
  description?: string | null
  run_count: number
  last_run_at?: string | null
  step_count: number
  tools: string[]
  parameters: SkillParameterView[]
}

export interface SkillCandidate {
  step_count: number
  tools: string[]
  request?: string | null
}

interface Props {
  open: boolean
  onClose: () => void
  /** 重放：把展开后的指令交给外层当普通消息发送 */
  onRun: (prompt: string) => void
}

export function SkillPanel({ open, onClose, onRun }: Props) {
  const [skills, setSkills] = useState<SkillView[]>([])
  const [candidate, setCandidate] = useState<SkillCandidate | null>(null)
  const [newName, setNewName] = useState('')
  const [saving, setSaving] = useState(false)
  const [error, setError] = useState('')
  // 正在填参数的技能；null 表示没有展开任何一个
  const [running, setRunning] = useState<SkillView | null>(null)
  const [argValues, setArgValues] = useState<Record<string, string>>({})
  // 云同步：登录后可用。未登录时按钮仍在，点了会提示先登录。
  const [syncMsg, setSyncMsg] = useState('')
  const [importCode, setImportCode] = useState('')

  const load = useCallback(async () => {
    try {
      const resp = await sendToHostWithResponse({ type: 'skill_list', payload: {} }, 'skill_list')
      if (!resp) return
      setSkills(resp.payload?.skills ?? [])
      setCandidate(resp.payload?.candidate ?? null)
    } catch {
      // 读不到技能列表不该阻塞面板
    }
  }, [])

  useEffect(() => {
    if (open) {
      setError('')
      setRunning(null)
      void load()
    }
  }, [open, load])

  const save = async () => {
    if (!newName.trim()) return setError('请给技能起个名字')
    setSaving(true)
    setError('')
    try {
      const resp = await sendToHostWithResponse(
        { type: 'skill_save', payload: { name: newName.trim() } },
        'skill_saved',
      )
      if (!resp) throw new Error('保存失败')
      setNewName('')
      setCandidate(null)
      await load()
    } catch (e: any) {
      setError(e?.message || '保存失败')
    } finally {
      setSaving(false)
    }
  }

  const remove = async (id: string) => {
    try {
      await sendToHostWithResponse({ type: 'skill_delete', payload: { id } }, 'skill_deleted')
      await load()
    } catch (e: any) {
      setError(e?.message || '删除失败')
    }
  }

  const sync = async () => {
    setSyncMsg('')
    setError('')
    try {
      const resp = await sendToHostWithResponse({ type: 'skill_sync', payload: {} }, 'skill_synced')
      if (!resp) throw new Error('同步失败')
      const { uploaded, redacted } = resp.payload
      setSyncMsg(
        redacted > 0
          ? `已同步 ${uploaded} 个技能。其中 ${redacted} 个含本地文件路径，上传前已移除。`
          : `已同步 ${uploaded} 个技能。`,
      )
    } catch (e: any) {
      setError(e?.message || '同步失败')
    }
  }

  const pull = async () => {
    setSyncMsg('')
    setError('')
    try {
      const resp = await sendToHostWithResponse({ type: 'skill_pull', payload: {} }, 'skill_pulled')
      if (!resp) throw new Error('拉取失败')
      setSyncMsg(`从云端恢复了 ${resp.payload.restored} 个技能。`)
      await load()
    } catch (e: any) {
      setError(e?.message || '拉取失败')
    }
  }

  const share = async (skill: SkillView) => {
    setSyncMsg('')
    setError('')
    try {
      const resp = await sendToHostWithResponse(
        { type: 'skill_share', payload: { id: skill.id } },
        'skill_shared',
      )
      if (!resp) throw new Error('分享失败')
      const { share_code, redacted } = resp.payload
      setSyncMsg(
        `分享码：${share_code}` +
          (redacted ? '（其中的本地文件路径已在上传前移除）' : ''),
      )
    } catch (e: any) {
      setError(e?.message || '分享失败')
    }
  }

  const importSkill = async () => {
    if (!importCode.trim()) return
    setSyncMsg('')
    setError('')
    try {
      const resp = await sendToHostWithResponse(
        { type: 'skill_import', payload: { share_code: importCode.trim() } },
        'skill_imported',
      )
      if (!resp) throw new Error('导入失败')
      setImportCode('')
      setSyncMsg(`已导入技能「${resp.payload.name}」。`)
      await load()
    } catch (e: any) {
      setError(e?.message || '导入失败')
    }
  }

  const beginRun = (skill: SkillView) => {
    setRunning(skill)
    const defaults: Record<string, string> = {}
    for (const parameter of skill.parameters) defaults[parameter.name] = parameter.default_value
    setArgValues(defaults)
  }

  const confirmRun = async () => {
    if (!running) return
    try {
      const resp = await sendToHostWithResponse(
        { type: 'skill_prepare', payload: { id: running.id, arguments: argValues } },
        'skill_prompt',
      )
      if (!resp) throw new Error('技能准备失败')
      onRun(resp.payload.prompt)
      onClose()
    } catch (e: any) {
      setError(e?.message || '运行失败')
    }
  }

  if (!open) return null

  return (
    <div className="config-overlay" onClick={onClose}>
      <div className="config-panel" onClick={(e) => e.stopPropagation()}>
        <div className="config-header">
          <h3>技能库</h3>
          <button className="config-close-btn" onClick={onClose} aria-label="关闭">
            ×
          </button>
        </div>

        <div className="config-body">
          {/* 刚刚成功完成的多步任务：趁用户还记得它做了什么时提议保存 */}
          {candidate && !running && (
            <div className="config-section skill-candidate">
              <p className="skill-candidate-title">
                刚完成的任务有 {candidate.step_count} 个步骤，保存下来下次一键重放？
              </p>
              {candidate.request && <p className="skill-candidate-req">{candidate.request}</p>}
              <div className="skill-candidate-row">
                <input
                  className="config-input"
                  type="text"
                  placeholder="给技能起个名字，例如：月度销售报表"
                  value={newName}
                  onChange={(e) => setNewName(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' && !saving) void save()
                  }}
                  disabled={saving}
                />
                <button className="config-save-btn" onClick={() => void save()} disabled={saving}>
                  {saving ? '保存中…' : '保存'}
                </button>
              </div>
            </div>
          )}

          {error && <p className="config-error">{error}</p>}
          {syncMsg && <p className="skill-sync-msg">{syncMsg}</p>}

          {!running && (
            <div className="config-section skill-sync">
              <div className="skill-sync-row">
                <button className="config-toggle-btn" onClick={() => void sync()}>
                  同步到云端
                </button>
                <button className="config-toggle-btn" onClick={() => void pull()}>
                  从云端恢复
                </button>
              </div>
              <div className="skill-sync-row">
                <input
                  className="config-input"
                  type="text"
                  placeholder="输入分享码导入他人的技能"
                  value={importCode}
                  onChange={(e) => setImportCode(e.target.value)}
                  onKeyDown={(e) => e.key === 'Enter' && void importSkill()}
                />
                <button className="config-toggle-btn" onClick={() => void importSkill()}>
                  导入
                </button>
              </div>
              <p className="config-apikey-hint">
                同步需要登录账号。上传前会移除本地文件路径与工作簿名——同步是备份，
                分享才是公开，两者分开。
              </p>
            </div>
          )}

          {running ? (
            <div className="config-section">
              <p className="skill-run-title">运行「{running.name}」</p>
              {running.parameters.length === 0 ? (
                <p className="config-apikey-hint">这个技能没有参数，直接运行即可。</p>
              ) : (
                running.parameters.map((parameter) => (
                  <label key={parameter.name} className="skill-param">
                    {parameter.label}
                    <input
                      className="config-input"
                      type="text"
                      value={argValues[parameter.name] ?? ''}
                      onChange={(e) =>
                        setArgValues((prev) => ({ ...prev, [parameter.name]: e.target.value }))
                      }
                    />
                  </label>
                ))
              )}
              <div className="skill-run-actions">
                <button className="config-toggle-btn" onClick={() => setRunning(null)}>
                  返回
                </button>
                <button className="config-save-btn" onClick={() => void confirmRun()}>
                  运行
                </button>
              </div>
              <p className="config-apikey-hint">
                运行时仍会像平常一样逐步确认，高风险操作照常显示变更预览。
              </p>
            </div>
          ) : skills.length === 0 ? (
            <div className="config-section">
              <p className="config-apikey-hint">
                还没有技能。完成一个多步任务后，回到这里就能把它保存下来重复使用。
              </p>
            </div>
          ) : (
            <div className="config-section">
              {skills.map((skill) => (
                <div key={skill.id} className="skill-item">
                  <div className="skill-item-main">
                    <span className="skill-item-name">{skill.name}</span>
                    <span className="skill-item-meta">
                      {skill.step_count} 步
                      {skill.run_count > 0 ? ` · 已运行 ${skill.run_count} 次` : ''}
                      {skill.parameters.length > 0 ? ` · ${skill.parameters.length} 个参数` : ''}
                    </span>
                  </div>
                  <div className="skill-item-actions">
                    <button className="config-toggle-btn" onClick={() => beginRun(skill)}>
                      运行
                    </button>
                    <button className="config-toggle-btn" onClick={() => void share(skill)}>
                      分享
                    </button>
                    <button className="config-toggle-btn" onClick={() => void remove(skill.id)}>
                      删除
                    </button>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
