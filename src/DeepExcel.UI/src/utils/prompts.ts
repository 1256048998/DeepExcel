// src/DeepExcel.UI/src/utils/prompts.ts
// 常用提示词模板：类型定义 + localStorage 读写
// ★ 纯前端本地存储，无需 C# 端持久化

export interface PromptTemplate {
  id: string          // UUID
  title: string       // 显示名称（如 "清洗数据"）
  content: string     // 提示词内容（如 "清洗A列数据，去除空格和特殊字符"）
  createdAt: number   // 创建时间戳
  source: 'history' | 'manual'  // 来源：历史消息保存 / 手动创建
}

const STORAGE_KEY = 'deepexcel_prompts'

/** 读取所有提示词 */
export function loadPrompts(): PromptTemplate[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    if (!raw) return []
    const arr = JSON.parse(raw)
    if (!Array.isArray(arr)) return []
    return arr
  } catch {
    return []
  }
}

/** 保存所有提示词（全量覆盖） */
export function savePrompts(prompts: PromptTemplate[]): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(prompts))
  } catch (e) {
    console.warn('[prompts] savePrompts failed:', e)
  }
}

/** 新增提示词，返回新数组 */
export function addPrompt(prompts: PromptTemplate[], title: string, content: string, source: 'history' | 'manual' = 'manual'): PromptTemplate[] {
  const newPrompt: PromptTemplate = {
    id: _genId(),
    title: title.trim() || '未命名提示词',
    content: content.trim(),
    createdAt: Date.now(),
    source,
  }
  const next = [newPrompt, ...prompts]
  savePrompts(next)
  return next
}

/** 更新提示词 */
export function updatePrompt(prompts: PromptTemplate[], id: string, updates: Partial<Pick<PromptTemplate, 'title' | 'content'>>): PromptTemplate[] {
  const next = prompts.map(p =>
    p.id === id ? { ...p, ...updates, title: (updates.title ?? p.title).trim() || p.title } : p
  )
  savePrompts(next)
  return next
}

/** 删除提示词 */
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

// ★ 简单 ID 生成（无需严格 UUID，前端本地用）
function _genId(): string {
  return 'p_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 8)
}
