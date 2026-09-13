import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'
import rehypeHighlight from 'rehype-highlight'
import 'highlight.js/styles/github.css'

interface Props {
  content: string
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
