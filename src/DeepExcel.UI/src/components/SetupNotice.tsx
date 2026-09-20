interface Props {
  onConfigure: () => void
}

/**
 * 「还没有可用的模型」引导。
 *
 * 装完打开面板，欢迎语会鼓励用户"试试说：在 A1 写入 =SUM(B1:B10)"。在配好
 * 供应商之前照做，消息会一路发到 sidecar，然后撞上一个来自模型 API 的技术性
 * 错误——用户既不知道哪里出了问题，也不知道下一步该干什么。**欢迎语等于在
 * 主动把人往坑里引。**
 *
 * 所以这条引导只在确实没有可用模型时出现，并且 sendMessage 会在同样的条件下
 * 拦住发送：光有提示挡不住直接按回车的人。
 */
export function SetupNotice({ onConfigure }: Props) {
  return (
    <div className="update-banner blocked" role="status">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor"
           strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
        <circle cx="12" cy="12" r="10" />
        <line x1="12" y1="8" x2="12" y2="12" />
        <line x1="12" y1="16" x2="12.01" y2="16" />
      </svg>
      <span className="update-banner-text">
        还没有可用的模型，先配置一个供应商
      </span>
      <button
        className="header-btn primary"
        onClick={onConfigure}
        title="填入任一供应商的 API Key 后即可开始"
        type="button"
      >
        去配置
      </button>
    </div>
  )
}
