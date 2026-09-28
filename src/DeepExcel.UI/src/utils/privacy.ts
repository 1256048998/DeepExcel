// 「数据与隐私」页的文案。每一句都对应一处实现，改实现要同步改这里：
//   - 发给模型：WorkbookSession.BuildContext 只带表名、当前表、选区地址和大小、附件名；
//     单元格是 AI 用工具按需读的（read_range 等），附件内容在你附上时才发
//   - 托管转发：server/app/proxy/router.py 只写 UsageRecord（模型、token 数、耗时、状态码），不存请求和回复
//   - 本机数据：%LOCALAPPDATA%\DeepExcel、%APPDATA%\DeepExcel；Key 和登录凭据用 DPAPI 加密
//   - 使用统计：TelemetryReporter 只在登录后发送；字段以服务端白名单为准（server/app/routers/telemetry.py）
//   - 技能同步：只在你点「同步 / 分享」时上传，SkillScrubber 先去掉文件路径和工作簿名
// 有测试守卫：privacy.test.ts 对照服务端白名单和客户端实际上报的事件

export const PRIVACY_SUMMARY =
  '工作簿内容只发给你选择的模型服务。DeepExcel 不保存单元格内容、你输入的文字和 AI 的回复。'

export const SENT_TO_MODEL: string[] = [
  '你发的消息，以及每轮开头的简要上下文：表名、当前表、选区的地址和大小、附件名（不含单元格的值）',
  'AI 为完成任务用工具读取的单元格——按需读取，不会整本上传',
  '你附上的文件内容（图片、PDF、文本），以及这个工作簿的记忆笔记',
]

export const ROUTE_TEXT = {
  byok: '自带 Key：请求从这台电脑直接发到你配置的模型服务商，不经过 DeepExcel 服务器。',
  hosted: '账号托管：请求经 DeepExcel 服务器转发给模型服务，服务器只记用量（模型、token 数、耗时），不保存请求和回复的内容。',
}

export const STORED_LOCALLY: string[] = [
  '对话历史、附件、工作簿记忆、写入前的备份快照、技能库和运行日志，都存在这台电脑的用户目录里',
  'API Key 和登录凭据用 Windows 加密（DPAPI）保存，只有你的 Windows 账户能解开',
]

export const TELEMETRY_INTRO =
  '只在登录账号后发送；没登录时不会离开这台电脑。服务器按公开的白名单接收，名单外的字段直接丢弃。'

// 每个事件、每个字段说的是什么。服务端白名单里的字段必须都在这里有说明（测试守卫）
export const TELEMETRY_EVENTS: Record<string, { label: string; fields: Record<string, string> }> = {
  session_start: {
    label: '打开面板',
    fields: {
      client_version: 'DeepExcel 版本',
      os_version: 'Windows 版本',
      office_version: 'Office 版本',
      office_bitness: 'Office 位数',
      host: 'Excel 还是 WPS',
    },
  },
  task_complete: {
    label: '一次任务结束',
    fields: {
      session_id: '随机会话编号',
      duration_ms: '耗时',
      turn_count: '步数',
      tool_sequence: '用过的工具名',
      tokens_in: '输入 token 数',
      tokens_out: '输出 token 数',
      provider: '模型服务商',
      model: '模型名',
      outcome: '结果（完成 / 出错 / 停止 / 追问）',
    },
  },
  tool_error: {
    label: '工具出错',
    fields: {
      tool_name: '工具名',
      error_code: '错误分类码（不含原始报错文字）',
    },
  },
  startup_error: {
    label: '启动失败',
    fields: {
      diagnostic_code: '诊断码',
      client_version: 'DeepExcel 版本',
    },
  },
  update_event: {
    label: '自动更新',
    fields: {
      phase: '阶段',
      outcome: '结果',
      reason_code: '失败原因码',
      from_version: '原版本',
      to_version: '新版本',
      duration_ms: '耗时',
    },
  },
}

// 服务端白名单里有、但客户端目前不发的事件（测试会确认客户端代码里确实没有上报它们）
export const TELEMETRY_NOT_SENT = ['user_feedback']

export const NEVER_COLLECTED: string[] = [
  '单元格内容',
  '工作簿名、文件名和路径',
  '你输入的文字',
  'AI 的回复',
  'API Key 和登录凭据',
  '原始报错文字',
]

export const SKILL_SYNC_TEXT =
  '技能只在你点「同步到云端」或分享技能时上传到账号，上传前会去掉文件路径和工作簿名。'
