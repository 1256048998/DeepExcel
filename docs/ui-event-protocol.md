# ui_event：侧车 → 面板的事件信封

侧车以前只发 `stream_delta` / `tool_use` / `stream_end`：面板只知道「调用了哪个工具」，
不知道它何时结束、成没成功、错在哪，只能显示一串英文工具名。`ui_event` 补上这些信息，
让面板能像 Claude Code 那样每步一行：

```
⏺ 读取 A1:C4                         0.1s
  ⎿ 4 行 × 3 列
⏺ 写入公式 D2 =SUM(A2:C2)             0.1s
  ⎿ D 列已被保护，不能写入
     改写到 E 列
⏺ 写入公式 E2 =SUM(A2:C2)             0.1s
  ⎿ 已写入 E2
完成 · 3 步（1 步失败） · 7.9s
```

## 信封

stdout 一行一个 JSON：

```json
{"type": "ui_event", "event": {"v": 1, "kind": "tool_start", "ts": 1758770000000, ...}}
```

- `v`：协议版本，当前 1。只加字段不改语义；要改语义就升版本。
- `ts`：侧车发出时的毫秒时间戳。
- 值为 `null` 的字段不发。

**宿主（C# `PythonSidecar`/`MessageBridge`、WPS `sidecar-host.js`）原样转发给面板**，
面板收到的是 `{type: "ui_event", payload: <event>}`。宿主唯一读取的是 `run_summary.outcome`
（C# 任务轨迹据此区分成功 / 出错 / 中断）。

## kind

| kind | 字段 | 含义 |
|---|---|---|
| `tool_start` | `id`, `name`, `args` | 模型发出一次工具调用。`id` 是 SDK 的 tool_use_id；`name` 不带 `mcp__excel__` 前缀；`args` 是显示用副本（长字符串截断到 4000 字，二维数组只留 `{__shape:[行,列], head:[前 3 行]}`） |
| `tool_end` | `id`, `name`, `ok`, `duration_ms`, `summary?`, `error?`, `check?` | 与同 `id` 的 `tool_start` 配对。`summary` 是一句话结果（「20 行 × 4 列」）；`error` 是 `{code, message, hint?}`；`check` 是写后自动体检 `{ok, summary}`（新增公式错误、新增外部链接），面板只在 `ok: false` 时显示；`checkpoint_id` 是这一步执行前单独存的检查点，面板据此显示「回到这一步之前」（发 `rollback_snapshot`） |
| `tool_gen` | `id`, `name`, `chars`, `lines?`, `preview?` | 模型还在生成这次调用的参数（每 250ms 最多一次）。代码类工具（execute_vba / execute_jsa / execute_python）带目前写到的代码 `preview`（最后 4000 字）；其他工具只报 `chars`。之后同一 `id` 的 `tool_start` 原地接替这一行 |
| `status` | `text`, `tool?` | 当前在做什么：等待用户确认、正在停止、看门狗（20 秒没动静「仍在等待模型响应」；某一步执行超过 15 秒「这一步执行中」；超过 120 秒提示可能被 Excel 对话框挡住、可以停止）。等用户确认 / 回答期间看门狗不催。`tool` 为 `explore_workbook` 时是分头摸底的进度（「分头摸底：1/3 个子任务完成（#2 find 应收）」），进度持续更新期间看门狗不发卡住提示。`text` 为空表示清除 |
| `plan` | `items: [{content, status}]` | 模型用 `todo_write` 维护的计划（status：pending / in_progress / completed，最多一条 in_progress）。面板在输入框上方显示计划胶囊；`todo_write` 本身不作为一步显示 |
| `plan_proposal` | `summary`, `steps: [{action, target?, detail?}]`, `risks: []` | 模型用 `present_plan` 提交的变更方案（「只出方案」模式下必须用它收尾）。面板显示方案卡片：批准并执行（每步确认）/ 批准并自动应用 / 继续修改；`present_plan` 本身不作为一步显示 |
| `compaction` | `trigger`, `pre_tokens?`, `prev_pct?`, `curr_pct?` | 上下文被压缩。`trigger` 为 `auto`/`manual`（CLI 的 compact_boundary）或 `detected`（没收到 compact_boundary、但上下文占比骤降超过 40%） |
| `error` | `code`, `message`, `hint`, `retryable`, `detail` | 整轮失败（API 报错、异常）。`message`/`hint` 是给用户的中文；`detail` 是原始报错前 500 字，只供诊断 |
| `steer_delivered` | `count` | 任务进行中用户发的插话已在某个工具结果之后交给模型（PostToolUse 的 additionalContext） |
| `steer_deferred` | `count` | 本轮结束前没有工具结果可以附带，插话转成下一条普通消息，侧车接着处理 |
| `run_summary` | `outcome`, `tool_calls`, `failed_calls`, `duration_ms`, `num_turns?`, `input_tokens?`, `output_tokens?` | 每轮结束、`stream_end` 之前恰好一次。`outcome`：`success` / `max_turns` / `error` / `interrupted` |

### 保证

- 每个 `tool_start` 都会有一个 `tool_end`。中断（`code: interrupted`）、出错（`aborted`）
  或本轮结束时仍没有结果（`no_result`）的调用，侧车会补发 `ok: false` 的 `tool_end`，
  面板上那一行不会一直转圈。
- `run_summary` 在 `stream_end` 之前发出，每轮最多一次。
- 被权限钩子拒绝的调用：`tool_end.ok = false`，`error.code = "denied"`。

### error.code

整轮错误（`kind: error`）：`auth`、`quota`、`task_limit`、`rate_limit`、`context_too_long`、
`model_not_found`、`timeout`、`network`、`cli_missing`、`unknown`。分类规则在
`src/DeepExcel.Sidecar/ui_events.py` 的 `_ERROR_RULES`。

引擎起不来（启动自检，`src/DeepExcel.Sidecar/selfcheck.py`）：`os_too_old`（Windows 早于 10 1809）、
`cli_missing`、`cli_blocked`（杀毒 / 应用管控拦截）、`cli_incompatible`、`cli_timeout`、`cli_crashed`、
`connect_failed`。侧车启动时立即发一次，此后每条用户消息都回同一个诊断并发 `stream_end`；
`retryable` 恒为 false，`detail` 里是原始错误。

工具错误（`tool_end.error`）：C# 结果里带 `error_code` 时用它，否则 `tool_failed`；
被拒绝为 `denied`；侧车补发的为 `interrupted` / `aborted` / `no_result`。

## 停止与插话（F9）

- **停止**：宿主发 `cancel`。侧车调用 `client.interrupt()` 让 CLI 停下当前回合，等它补发工具结果和
  ResultMessage（最多 15 秒，期间发 `status: 正在停止…`），面板上的步骤正常收尾，`run_summary.outcome =
  interrupted`。等不到就硬切，并在下一轮开始前把 CLI 的残留输出读到上一轮的 ResultMessage 为止——
  否则下一个问题会先读到上一轮剩下的回答。等待中的工具调用、权限确认、澄清问题收到停止后立即返回。
- **插话**：任务进行中面板照常可以输入，发出的 `user_message` 带 `steer: true`。侧车在下一个工具结果之后
  把它作为 `<user-interjection>` 交给模型（`steer_delivered`）；本轮结束还没送达就作为下一条消息处理
  （`steer_deferred`）。按了停止则插话一起作废。和插话同一批已经发出的工具调用撤不回来（Claude Code 也一样）。
- C# 对插话不开新回合（不重置写前备份、不新建任务轨迹）；侧车只在确实有一轮在跑时才把 `steer` 当插话，
  空闲时当普通消息。

## 权限模式

面板 → 宿主 → 侧车，宿主只转发、不保存；三种取值 `default`（每步确认）/ `accept_writes`（本次会话自动应用写入）/ `plan`（只出方案），见 `src/DeepExcel.Sidecar/permission_modes.py`。

- 每条 `user_message` 带 `permission_mode`：面板是唯一来源，侧车重启回到默认时也能对上。
- 任务进行中切换发 `set_permission_mode {mode}`，侧车下一次工具调用就按新模式判断。
- `accept_writes`：批量写入、清洗不再弹确认（每步照常有检查点）；删除、清空、回滚、执行代码仍然确认。
- `plan`：只读工具、`clarify_intent`、`todo_write`、`load_skill`、`update_workbook_notes`、`present_plan` 之外一律在 PreToolUse 拒绝，`tool_end.error.code = "denied"`。
- 模式不写进任何配置；面板重开回到 `default`。

## 副本试跑（Excel）

`execute_vba` 要确认时，宿主先 SaveCopyAs 出副本，在隐藏的另一个 Excel 进程（DCOM 新实例，核对 PID）里跑同一段代码，再弹确认。代码见 `src/DeepExcel.AddIn/Executor/LabRunner.cs`、`src/DeepExcel.AddIn/Bridge/LabTrialHandlers.cs`。

- 试跑期间发 `ui_event {kind: status, text: "正在工作簿副本上试跑这段代码…"}`，出结果后发空 `text` 清掉。
- `permission_request.preview.trial`：`success`、`error`、`timed_out`、`duration_ms`、`errors_before` / `errors_after`（公式错误格数）、`note`（执行器提示，含自动点掉的弹窗）、`not_representative[]`（外部链接、数据连接、自带宏时事件被关、代码读路径 / 其他工作簿 / 时间随机数）、`stale`。`changes` 是副本前后的公式文本差异。没能试跑（副本起不来、csv 等格式）时没有 `trial`，`previewable=false`，`reason` 带上原因。
- 用户之后改到的格和试跑改动的区域相交：宿主发 `permission_preview_stale {request_id, message}`，面板标出过期并给「重新试跑」，点了发 `rerun_trial {request_id}`；宿主重新存副本再跑，用同一个 `request_id` 再发一次 `permission_request`。
- 允许后在真实工作簿上执行的是同一段代码（照常快照），副本从不拷回。超时 20 秒结束副本进程；副本进程按 PID + 创建时间登记在 `%LOCALAPPDATA%\DeepExcel\lab\processes.json`，主人进程不在了就在下次启动或试跑前回收。副本实例里插件保持被动（`ThisAddIn.IsAutomationInstance`：UserControl=false 且不可见），不建面板、不起侧车。
- WPS 的 `execute_jsa` 没有副本试跑。

## 与旧消息的关系

- `tool_use` 仍然发：C# 用它记对话历史和任务轨迹，WPS 用它记对话历史。面板不再渲染它。
- 旧的 `compacted` 消息已删除（C# 从来没解析过它，Excel 里压缩提示从未出现过）。
- API 报错以前作为助手文本流给用户（英文原文），现在是 `kind: error`。

## 面板

- 事件 → 消息列表：`src/DeepExcel.UI/src/utils/uiEvents.ts` 的 `applyUiEvent`（纯函数，有测试）。
- 工具的中文叙事：`src/DeepExcel.UI/src/utils/toolCatalog.ts`。`toolCatalog` 的测试会读侧车
  `excel_tools.py` 的注册表逐条比对，新增工具没写文案会失败。

## 验证

- 单测：`src/DeepExcel.Sidecar/tests/test_ui_events.py`、`src/DeepExcel.UI/src/utils/uiEvents.test.ts`、
  `TelemetryReporterTests.RunSummaryOutcomeDecidesTheTraceOutcome`。
- 真实模型：`python scripts/live_sidecar_events.py` —— 按 WPS 的方式启动侧车（自己读本机
  DeepExcel 配置和 DPAPI 凭据，脚本碰不到 Key），假装宿主应答工具调用，并故意让一次写入失败，
  检查事件配对、失败上报和终态行。`--interrupt`：停止后再问新问题，回答不能是上一轮的残留；
  `--steer`：工具执行期间插话「改写到 F2」，模型应当照做。每个场景消耗几千 token。
