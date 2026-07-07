/**
 * 厂商图标：圆形品牌色背景 + 白色首字母
 * 避免 AI 风格竖线，使用真实品牌色
 */

interface IconProps {
  size?: number
}

const CircleIcon = ({ color, letter, size = 28 }: IconProps & { color: string; letter: string }) => (
  <svg width={size} height={size} viewBox="0 0 28 28" xmlns="http://www.w3.org/2000/svg">
    <circle cx="14" cy="14" r="13" fill={color} />
    <text
      x="14"
      y="14"
      textAnchor="middle"
      dominantBaseline="central"
      fill="white"
      fontSize="13"
      fontWeight="600"
      fontFamily="-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif"
    >
      {letter}
    </text>
  </svg>
)

export const providerIcons: Record<string, (props: IconProps) => JSX.Element> = {
  anthropic: (props) => <CircleIcon color="#D97757" letter="A" {...props} />,
  deepseek: (props) => <CircleIcon color="#4D6BFE" letter="D" {...props} />,
  stepfun: (props) => <CircleIcon color="#7B61FF" letter="S" {...props} />,
  openai: (props) => <CircleIcon color="#000000" letter="O" {...props} />,
  kimi: (props) => <CircleIcon color="#1A1A1A" letter="K" {...props} />,
  qwen: (props) => <CircleIcon color="#615CED" letter="Q" {...props} />,
  zhipu: (props) => <CircleIcon color="#3B82F6" letter="Z" {...props} />,
  minimax: (props) => <CircleIcon color="#2563EB" letter="M" {...props} />,
  doubao: (props) => <CircleIcon color="#FF1A4B" letter="D" {...props} />,
  custom: (props) => <CircleIcon color="#6B7280" letter="+" {...props} />
}

/** 厂商显示顺序（左侧列表按此顺序渲染） */
export const providerOrder = [
  'anthropic', 'deepseek', 'stepfun', 'openai',
  'kimi', 'qwen', 'zhipu', 'minimax', 'doubao', 'custom'
]
