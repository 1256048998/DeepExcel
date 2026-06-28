interface Props {
  value: string
  onChange: (val: string) => void
  onSend: () => void
  onStop?: () => void
  disabled: boolean
  isClarifying?: boolean
}

export function InputArea({ value, onChange, onSend, onStop, disabled, isClarifying }: Props) {
  return (
    <div className="input-area">
      {isClarifying && (
        <div className="clarify-hint">
          💡 请回答上面的澄清问题
        </div>
      )}
      <div className="input-row">
        <textarea
          value={value}
          onChange={e => onChange(e.target.value)}
          onKeyDown={e => {
            if (e.key === 'Enter' && !e.shiftKey) {
              e.preventDefault()
              onSend()
            }
          }}
          placeholder={isClarifying ? '输入你的回答...' : '描述你的Excel任务...'}
          rows={2}
          disabled={disabled}
        />
        {disabled && onStop ? (
          <button
            onClick={onStop}
            className="stop-button"
          >
            ⏹ 停止
          </button>
        ) : (
          <button
            onClick={onSend}
            disabled={disabled || !value.trim()}
            className="send-button"
          >
            发送
          </button>
        )}
      </div>
    </div>
  )
}
