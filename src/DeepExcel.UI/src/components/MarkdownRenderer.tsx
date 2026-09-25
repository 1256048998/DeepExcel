import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'
import rehypeHighlight from 'rehype-highlight'
import 'highlight.js/styles/github-dark.css'
import { Children, isValidElement } from 'react'
import type { ReactElement, ReactNode } from 'react'
import { CopyButton } from './CopyButton'

interface Props {
  content: string
}

const LANG_LABELS: Record<string, string> = {
  vba: 'VBA', vb: 'VBA', vbnet: 'VBA', excel: '公式', formula: '公式',
  js: 'JavaScript', javascript: 'JavaScript', ts: 'TypeScript', py: 'Python', python: 'Python',
  sql: 'SQL', json: 'JSON', bash: 'Shell', sh: 'Shell', powershell: 'PowerShell', csv: 'CSV',
}

// rehype-highlight 把代码拆成一堆 span；复制时要拼回纯文本
function textOf(node: ReactNode): string {
  if (node == null || typeof node === 'boolean') return ''
  if (typeof node === 'string' || typeof node === 'number') return String(node)
  if (Array.isArray(node)) return node.map(textOf).join('')
  if (isValidElement(node)) return textOf((node.props as any).children)
  return ''
}

export function MarkdownRenderer({ content }: Props) {
  return (
    // ★ .md-body 关掉从 .message 继承的 white-space: pre-wrap。
    // markdown 已经把换行解析成块级元素了，再叠一层 pre-wrap 会把段内的软换行
    // 渲染成硬换行，和块级 margin 撞在一起就是用户反馈的"空行过多"。
    // 纯文本消息不走这里，仍由 .message 的 pre-wrap 保留换行。
    <div className="md-body">
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        rehypePlugins={[rehypeHighlight]}
        components={{
          // 代码块
          code: ({ node, className, children, ...props }) => {
            const isInline = !className
            if (isInline) {
              return (
                <code className="md-inline-code" {...props}>
                  {children}
                </code>
              )
            }
            return (
              <code className={className} {...props}>
                {children}
              </code>
            )
          },
          // 代码块外壳：深色卡片 + 标题栏（语言 + 复制）
          pre: ({ children }) => {
            const code = Children.toArray(children).find(isValidElement) as ReactElement<any> | undefined
            const lang = /language-([\w+-]+)/.exec(code?.props?.className ?? '')?.[1] ?? ''
            return (
              <div className="md-code">
                <div className="md-code-header">
                  <span>{LANG_LABELS[lang.toLowerCase()] ?? (lang || '代码')}</span>
                  <CopyButton content={textOf(code?.props?.children).replace(/\n$/, '')} />
                </div>
                <pre>{children}</pre>
              </div>
            )
          },
          // 链接
          a: ({ href, children }) => (
            <a href={href} target="_blank" rel="noopener noreferrer">
              {children}
            </a>
          ),
          // 表格
          table: ({ children }) => (
            <div className="md-table-wrapper">
              <table className="md-table">{children}</table>
            </div>
          ),
          // 任务列表
          input: ({ type, checked, ...props }) => (
            <input type="checkbox" checked={checked} readOnly {...props} />
          ),
        }}
      >
        {content}
      </ReactMarkdown>
    </div>
  )
}
