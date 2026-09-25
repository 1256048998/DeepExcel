/**
 * 从 AI 回复里找「请用户选一个」的选项，渲染成可点击卡片。
 *
 * 以前任何编号列表都会被当成选项：「已完成：1. C 列… 2. D 列…」这样的总结也冒出
 * 「选择方案」按钮，按钮上还带着 **粗体** 的 markdown 原文。现在：
 *   - 明确标了选项的（方案 A / 选项 1 / A. / (A) / 【A】）照旧识别
 *   - 普通编号列表（1. / 一、）只有在它前一行是在让用户选（「选哪种」「有两种方案」「？」）时才算
 *   - 选项必须是连续的几行（中间可以空行），取第一组 ≥ 2 个的
 *   - 显示和发送的文字去掉 markdown 标记
 */

export interface Choice {
  /** 选项标签，如 "A" / "方案 A" / "1" */
  label: string
  /** 选项描述文本（点击后发送的内容） */
  value: string
}

const EXPLICIT: RegExp[] = [
  /^(方案|选项|Plan|Option)\s*([A-Z0-9一二三四五六七八九十])\s*[.、:：]\s*(.+)$/i,
  /^([A-Z])\s*[.、:：）)]\s*(.+)$/,
  /^\(([A-Z])\)\s*(.+)$/,
  /^【([A-Z一-龥])】\s*(.+)$/,
]

const NUMBERED: RegExp[] = [
  /^(\d+)\s*[.、:：）)]\s*(.+)$/,
  /^\((\d+)\)\s*(.+)$/,
  /^([一二三四五六七八九十])\s*[、.]\s*(.+)$/,
]

/** 前一行在让用户做选择 */
const CHOICE_CUE = /选择|选一|选哪|哪种|哪个|哪一|几种|两种|三种|方案|可以这样|你想|您想|你希望|您希望|还是|[？?]\s*$/

const MAX_VALUE_LENGTH = 120

export function stripMarkdown(text: string): string {
  return text
    .replace(/\[([^\]]+)\]\([^)]*\)/g, '$1')
    .replace(/(\*\*|__)(.+?)\1/g, '$2')
    .replace(/(\*|_)(.+?)\1/g, '$2')
    .replace(/`([^`]+)`/g, '$1')
    .trim()
}

function matchLine(line: string, patterns: RegExp[]): Choice | null {
  const trimmed = line.trim().replace(/^[-*+]\s+/, '')
  for (const re of patterns) {
    const m = trimmed.match(re)
    if (!m) continue
    const withPrefix = m.length === 4 && m[3] !== undefined
    const label = withPrefix ? `${m[1]} ${m[2]}` : m[1]
    const value = stripMarkdown(withPrefix ? m[3] : m[2])
    if (!value || value.length > MAX_VALUE_LENGTH) return null
    return { label, value }
  }
  return null
}

export function detectChoices(content: string): Choice[] {
  if (!content || content.length < 4) return []
  const lines = content.split('\n')

  for (let i = 0; i < lines.length; i++) {
    const explicit = matchLine(lines[i], EXPLICIT)
    const numbered = explicit ? null : matchLine(lines[i], NUMBERED)
    if (!explicit && !numbered) continue

    // 普通编号列表：看它上面最近的一行是不是在让用户选
    if (!explicit) {
      let prev = i - 1
      while (prev >= 0 && !lines[prev].trim()) prev--
      if (prev < 0 || !CHOICE_CUE.test(lines[prev].trim())) {
        // 跳过这一整组编号，免得组内第二项被当成新的开头
        while (i + 1 < lines.length && (matchLine(lines[i + 1], NUMBERED) || !lines[i + 1].trim())) i++
        continue
      }
    }

    const patterns = explicit ? EXPLICIT : NUMBERED
    const group: Choice[] = []
    let j = i
    for (; j < lines.length; j++) {
      if (!lines[j].trim()) continue
      const choice = matchLine(lines[j], patterns)
      if (!choice) break
      if (!group.some(c => c.label === choice.label)) group.push(choice)
    }
    if (group.length >= 2) return group
    i = j - 1
  }
  return []
}
