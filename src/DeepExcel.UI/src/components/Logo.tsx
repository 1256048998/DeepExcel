// DeepExcel 标志：品牌蓝圆角方块，里面是表格的三个格子，第四格换成 AI 星芒。
// 纯 SVG、不依赖外部资源，Excel 与 WPS 面板里都一样清晰。
export function LogoMark({ size = 22 }: { size?: number }) {
  return (
    <svg className="logo-mark" width={size} height={size} viewBox="0 0 24 24" aria-hidden="true">
      <defs>
        <linearGradient id="de-logo-bg" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0" stopColor="#3b82f6" />
          <stop offset="1" stopColor="#1a55c4" />
        </linearGradient>
      </defs>
      <rect width="24" height="24" rx="6.5" fill="url(#de-logo-bg)" />
      <rect x="5" y="5" width="6" height="6" rx="1.6" fill="#fff" />
      <rect x="13" y="5" width="6" height="6" rx="1.6" fill="#fff" fillOpacity="0.62" />
      <rect x="5" y="13" width="6" height="6" rx="1.6" fill="#fff" fillOpacity="0.62" />
      <path d="M16 12.2c.35 2.1 1.2 3.05 3.8 3.8-2.6.75-3.45 1.7-3.8 3.8-.35-2.1-1.2-3.05-3.8-3.8 2.6-.75 3.45-1.7 3.8-3.8z" fill="#fff" />
    </svg>
  )
}

export function Brand() {
  return (
    <div className="app-brand" title="DeepExcel">
      <LogoMark />
      <span className="app-brand-name">DeepExcel</span>
    </div>
  )
}
