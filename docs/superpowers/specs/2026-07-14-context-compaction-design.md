# 会话内自动压缩（Context Compaction）

**日期：** 2026-07-14
**状态：** 设计已完成，待实现

## 背景与动机

DeepExcel 使用 Claude Agent SDK（v0.2.109）+ Claude Code CLI（v2.1.190）作为 AI 底层。
当对话变长时，context window 的 token 会接近上限，导致 API 报错或对话中断。

Claude Code CLI 内置 **autocompact** 机制：当 token 接近上限时，自动把前面的消息摘要化，
以 `isCompactSummary` 标记存入会话，后续对话基于压缩后的上下文继续。

目前 DeepExcel 的 sidecar 没有感知这个机制——压缩发生时用户完全无感，突然看到 AI 的
回复"忘记"了前面说过的细节，会困惑。本设计的目标是：**当压缩发生时，在对话中插入一条
提示卡，让用户知道发生了什么。**

## 范围

- **只做会话内自动压缩的"可视化反馈"**——依赖 SDK/CLI 原生 autocompact，不自建压缩逻辑
- **不做 token 进度条**——Excel 一般用户使用不太会超上下文，常驻进度条是噪音
- **不做自定义压缩阈值**——用 CLI 默认值，YAGNI
- **不做跨会话摘要 / 长期记忆**——那是另一个功能

## 架构

```
┌──────────────────────────────────────────────────┐
│ Excel 面板（前端 React）                          │
│  消息列表                                          │
│   ├── ...正常消息...                              │
│   ├── ┌──────────────────────────────────────┐   │
│   │   │ ✓ 对话已自动压缩，保留了关键上下文    │   │  ← 压缩提示卡
│   │   └──────────────────────────────────────┘   │     （仅在压缩发生时出现）
│   └── ...压缩后的新消息...                       │
└──────────────────────────────────────────────────┘
          ▲ compacted 消息
          │
┌─────────┴────────────────────────────────────────┐
│ sidecar.py                                        │
│  handle_sdk_message:                              │
│    ├── StreamEvent → stream_delta（正常）         │
│    ├── AssistantMessage → tool_use（正常）        │
│    ├── SystemMessage → ★ 检测压缩事件            │
│    │     └── subtype 含 compact → 发 compacted   │
│    └── ResultMessage → stream_end（正常）         │
│                                                    │
│  ClaudeAgentOptions:                              │
│    settings={"autoCompactEnabled": true}          │
└──────────────────────────────────────────────────┘
          ▲ 控制协议
          │
┌─────────┴────────────────────────────────────────┐
│ Claude Code CLI 2.1.190（内置 autocompact）       │
│  token 接近上限 → 自动摘要化前面消息              │
│  发 SystemMessage(subtype=compact) 给 sidecar     │
└──────────────────────────────────────────────────┘
```

三层职责：
- **CLI**：执行压缩（摘要生成 + 上下文替换），发 SystemMessage 通知
- **sidecar**：检测 SystemMessage 中的压缩事件，转发 `compacted` 消息给前端
- **前端**：收到 `compacted` 时插入提示卡，平时无任何 UI 改动

## 组件设计

### A. sidecar.py 改动

**1. 显式启用 autocompact**

在 `ClaudeAgentOptions` 中通过 `settings` 传 `autoCompactEnabled: true`。CLI 默认启用，
但显式确认避免未来版本变更默认值时出问题。

```python
options = ClaudeAgentOptions(
    # ...existing options...
    settings=json.dumps({"autoCompactEnabled": True}),
)
```

**2. handle_sdk_message 增加 SystemMessage 处理**

当前 `handle_sdk_message` 只处理 `StreamEvent`、`AssistantMessage`、`ResultMessage`。
新增 `SystemMessage` 处理：检查 `subtype` 和 `data` 中是否含 compact 相关标记。

```python
elif isinstance(response, SystemMessage):
    # ★ 检测 autocompact 事件
    subtype = response.subtype or ""
    data = response.data or {}
    # CLI 发送的压缩事件 subtype 可能是 "compact" 或 data 里含 isCompactSummary
    if "compact" in subtype.lower() or data.get("isCompactSummary"):
        await write_message({"type": "compacted"})
```

> **注：** 压缩事件的确切 subtype 名称和 data 结构需要运行时验证。sidecar 会同时
> 检查 subtype 含 "compact" 和 data 含 "isCompactSummary" 两个条件，覆盖可能的变体。
> 同时加 stderr 日志记录所有 SystemMessage 的 subtype，便于后续确认。

**3. 兼容性处理**

`handle_sdk_message` 需要导入 `SystemMessage`。当前已从 `claude_agent_sdk.types` 导入了
`AssistantMessage`、`ResultMessage`、`StreamEvent`、`TextBlock`、`ToolUseBlock`，追加 `SystemMessage`。

### B. 前端改动（App.tsx + MessageList）

**1. App.tsx 增加 compacted 消息处理**

在 `onHostMessage` 的消息分发中，新增 `compacted` 类型处理：插入一条特殊的 assistant 消息。

```typescript
} else if (data.type === 'compacted') {
  // ★ 插入压缩提示卡（作为特殊 assistant 消息）
  setMessages(prev => [...prev, {
    role: 'assistant',
    content: '',
    type: 'compacted'  // 新消息类型，MessageList 用它渲染提示卡
  }])
}
```

**2. types.ts 扩展 Message 类型**

`Message` 接口增加 `'compacted'` 作为可选 `type` 值。

**3. MessageList.tsx 渲染压缩提示卡**

当 `message.type === 'compacted'` 时，渲染为浅色提示卡而非普通气泡：

```tsx
{msg.type === 'compacted' ? (
  <div className="message compacted-hint">
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" 
         stroke="currentColor" strokeWidth="2" strokeLinecap="round" 
         strokeLinejoin="round">
      <path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/>
      <path d="M21 3v5h-5"/>
    </svg>
    <span>对话已自动压缩，保留了关键上下文</span>
  </div>
) : (
  // 正常消息渲染
)}
```

### C. styles.css 新增样式

```css
/* 压缩提示卡：浅灰背景，紧凑，居中 */
.message.compacted-hint {
  display: flex;
  align-items: center;
  gap: 6px;
  margin: 8px auto;
  padding: 6px 12px;
  background: #f3f4f6;
  border: 1px solid #e5e7eb;
  border-radius: 6px;
  color: #6b7280;
  font-size: 12px;
  max-width: 90%;
  text-align: center;
}
```

遵循项目 UI 规范：浅灰背景（#f3f4f6）+ 灰色文字（#6b7280），不用 AI 风格蓝紫色。

## 数据流

```
1. 用户持续对话 → token 逐渐增长
2. CLI 检测 token 接近上限
3. CLI 执行 autocompact：
   a. 调用 Claude 生成前面消息的摘要
   b. 用摘要替换原始消息（isCompactSummary 标记）
   c. 发 SystemMessage(subtype=compact) 给 sidecar
4. sidecar handle_sdk_message 收到 SystemMessage
   → 检测到压缩事件
   → 发 {"type": "compacted"} 给前端
5. 前端 onHostMessage 收到 compacted
   → 插入 type='compacted' 的 assistant 消息
   → MessageList 渲染提示卡
6. CLI 继续处理用户消息（基于压缩后的上下文）
   → 正常 stream_delta / stream_end
```

## 错误处理

- **SystemMessage 解析失败**：sidecar 已有 try/except 包裹，失败时仅写 stderr 日志，不影响主流程
- **前端收到未知消息类型**：现有 `onHostMessage` 的 if/else 链忽略未知类型，不会崩溃
- **autocompact 未启用**（如 CLI 版本不支持）：不会发 SystemMessage，前端永远不显示提示卡，无副作用

## 测试策略

1. **手动测试**：构造超长对话（重复发送大段文本），触发 autocompact，确认提示卡出现
2. **日志验证**：sidecar stderr 记录所有 SystemMessage 的 subtype，确认压缩事件的实际 subtype 名称
3. **回归测试**：正常短对话不出现提示卡，功能不受影响

## 改动文件清单

| 文件 | 改动 |
|------|------|
| `src/DeepExcel.Sidecar/sidecar.py` | 导入 SystemMessage；handle_sdk_message 增加压缩检测；options 加 settings |
| `src/DeepExcel.UI/src/App.tsx` | onHostMessage 增加 compacted 处理 |
| `src/DeepExcel.UI/src/types.ts` | Message 接口增加 'compacted' 类型 |
| `src/DeepExcel.UI/src/components/MessageList.tsx` | 渲染压缩提示卡 |
| `src/DeepExcel.UI/src/styles.css` | .compacted-hint 样式 |

## 不做的事

- ❌ token 进度条（Excel 用户低频场景，常驻 UI 是噪音）
- ❌ 自定义压缩阈值（YAGNI，CLI 默认值够用）
- ❌ sidecar 层 token 计数（SDK 的 get_context_usage 是权威，自己数不准）
- ❌ 跨会话摘要 / 长期记忆（另一个功能）
- ❌ 压缩时保留某些消息不摘要（CLI 不支持自定义压缩策略）
