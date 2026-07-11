// src/DeepExcel.UI/src/utils/prompts.ts
// 常用提示词 / 技能：类型定义 + localStorage 读写
// ★ 纯前端本地存储（用户级），无需 C# 端持久化，跨工作簿保留

export type PromptType = 'prompt' | 'skill'

export interface PromptTemplate {
  id: string          // UUID
  title: string       // 显示名称（如 "清洗数据"）
  content: string     // 提示词内容（prompt=纯文本；skill=SKILL.md 格式，含 YAML frontmatter）
  type: PromptType    // 类型：prompt=普通提示词 / skill=Claude 标准技能
  createdAt: number   // 创建时间戳
  source: 'history' | 'manual'  // 来源：历史消息保存 / 手动创建
}

const STORAGE_KEY = 'deepexcel_prompts'

/** 读取所有提示词/技能 */
export function loadPrompts(): PromptTemplate[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (!raw) return []
    const arr = JSON.parse(raw)
    if (!Array.isArray(arr)) return []
    // ★ 兼容旧数据：无 type 字段时默认为 prompt
    return arr.map((p: any) => ({ ...p, type: (p.type as PromptType) || 'prompt' }))
  } catch {
    return []
  }
}

/** 保存所有提示词/技能（全量覆盖） */
export function savePrompts(prompts: PromptTemplate[]): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(prompts))
  } catch (e) {
    console.warn('[prompts] savePrompts failed:', e)
  }
}

/** 新增提示词/技能，返回新数组 */
export function addPrompt(
  prompts: PromptTemplate[],
  title: string,
  content: string,
  type: PromptType = 'prompt',
  source: 'history' | 'manual' = 'manual'
): PromptTemplate[] {
  const newPrompt: PromptTemplate = {
    id: _genId(),
    title: title.trim() || (type === 'skill' ? '未命名技能' : '未命名提示词'),
    content: content.trim(),
    type,
    createdAt: Date.now(),
    source,
  }
  const next = [newPrompt, ...prompts]
  savePrompts(next)
  return next
}

/** 更新提示词/技能 */
export function updatePrompt(prompts: PromptTemplate[], id: string, updates: Partial<Pick<PromptTemplate, 'title' | 'content' | 'type'>>): PromptTemplate[] {
  const next = prompts.map(p =>
    p.id === id ? { ...p, ...updates, title: (updates.title ?? p.title).trim() || p.title } : p
  )
  savePrompts(next)
  return next
}

/** 删除提示词/技能 */
export function deletePrompt(prompts: PromptTemplate[], id: string): PromptTemplate[] {
  const next = prompts.filter(p => p.id !== id)
  savePrompts(next)
  return next
}

/** 模糊匹配：title 和 content 都参与匹配 */
export function matchPrompts(prompts: PromptTemplate[], query: string): PromptTemplate[] {
  const q = query.trim().toLowerCase()
  if (!q) return prompts
  return prompts.filter(p =>
    p.title.toLowerCase().includes(q) ||
    p.content.toLowerCase().includes(q)
  )
}

/** 生成 SKILL.md 标准模板（遵循 Agent Skills 规范：YAML frontmatter + Markdown 指令） */
export function makeSkillTemplate(name?: string): string {
  const skillName = (name || 'my-skill').trim().toLowerCase().replace(/\s+/g, '-')
  return `---
name: ${skillName}
description: 在这里描述这个技能的用途（AI 会根据描述判断何时调用）
---

# 技能指令

在这里编写技能的具体指令。Claude 会按照这里的指令执行任务。

## 示例

当用户调用此技能时，按照以下步骤操作：
1. 第一步...
2. 第二一步...
3. ...
`
}

// ★ 简单 ID 生成（无需严格 UUID，前端本地用）
function _genId(): string {
  return 'p_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 8)
}
