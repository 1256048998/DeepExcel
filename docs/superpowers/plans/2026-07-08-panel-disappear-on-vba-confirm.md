# 专题：VBA/Python 确认后面板消失问题

## 一、问题现象

用户反馈：**经常在运行 VBA 或 Python 点击确认时，对话面板自动消失，且点"打开面板"也打不开。此时 AI 还在后台运行、更新 Excel，但用户看不见。**

## 二、日志证据（2026-07-08 测试3.xlsx 会话）

### 关键时序

| 时间 | 事件 | 来源日志 |
|---|---|---|
| 07:20:54 | AI 调用 execute_vba | deepexcel-20260708.log |
| 07:21:16 | AI 调用 rollback（用户点"是"确认） | deepexcel-20260708.log |
| 07:21:18 | **WorkbookBeforeClose: 测试3.xlsx 触发** | DeepExcel_Load.log |
| 07:21:18 | **SKIPPED (tool execution guard active)** | DeepExcel_Load.log |
| 07:21:21 | stream_end（任务结束，guard 恢复 false） | deepexcel-20260708.log |
| 07:21:25 | 用户点"打开面板" → **already visible** | DeepExcel_Load.log |
| 07:21:37/41/47 | 用户 3 次点"打开面板" → **already visible** | DeepExcel_Load.log |
| 07:21:48 | **Excel HostShutdown**（加载项全部关闭） | DeepExcel_Load.log |
| 07:21:51 | Removed CTP / OnBeginShutdown / OnDisconnection | DeepExcel_Load.log |

### 根因链
1. `PromptUserApproval` 用 `MessageBox.Show()` 同步阻塞 UI 线程
2. VBA/rollback 执行期间阻塞 → Excel 可能触发 `WorkbookBeforeClose`
3. `ExecutionGuard` 跳过关闭处理 → CTP/sidecar 状态不一致
4. 状态不一致累积 → Excel 异常退出（HostShutdown）
5. Excel 退出 → CTP 失效，字典保留失效对象，Visible 返回缓存值 true → "already visible 但看不见"

## 三、最优方案：PreToolUse Hook + 面板内抽屉式确认

### 设计理念（对标 Claude Code / Trae / Codex）

- **不阻塞 UI 线程**：用 SDK 的 PreToolUse hook 异步回调，替代 `MessageBox.Show()`
- **面板内抽屉**：从输入框上方"抽出"抽屉（slide-up 动画），在面板内显示确认卡片
- **保持 Excel 响应**：等待用户确认期间 Excel 正常响应，不触发误关闭

### 为什么用 PreToolUse hook 而非 can_use_tool

查 SDK 源码（[types.py:1803-1812](file:///C:/Users/qinju/AppData/Local/hermes/hermes-agent/venv/Lib/site-packages/claude_agent_sdk/types.py#L1803-L1812)）：
- `can_use_tool` 只在 CLI permission rules 评估为 "ask" 时触发
- 我们的工具都在 `allowed_tools` 里，**永远不会触发 can_use_tool**
- **PreToolUse hook** 对每个工具调用都触发，可返回 `permissionDecision: "allow"/"deny"`

### 完整数据流

```
AI 决定调用 execute_vba
  ↓
SDK 触发 PreToolUse hook（异步回调）
  ↓
sidecar.py hook 函数：检查 tool_name 是否高风险
  ├─ 低风险（read_range/write_value 等）→ 直接返回 allow
  └─ 高风险（execute_vba/execute_python/rollback/clean_data/remove_duplicates）
      → 发送 {type: "permission_request", tool, args, request_id} 到 C#
      → C# 转发到 WebView
      → 前端抽屉从输入框上方滑出：
        ┌─────────────────────────────┐
        │ ⚠ AI 想执行 VBA 代码         │
        │ 参数预览：...                │
        │ [拒绝]  [允许此次]           │
        └─────────────────────────────┘
      → 用户点击 → 前端发回 {type: "permission_response", request_id, decision}
      → C# → sidecar → hook 返回 permissionDecision
  ↓
SDK 根据 permissionDecision 执行或跳过工具
  ↓
工具执行（此时 UI 线程不阻塞，Excel 正常响应）
```

## 四、实施计划

### 任务 1：扩展 SidecarProtocol 协议（5 min）

**文件**：[SidecarProtocol.cs](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Sidecar/SidecarProtocol.cs)

新增两个消息类型：
```csharp
// Python → C#（请求权限）
public const string TypePermissionRequest = "permission_request";
// C# → Python（权限响应）
public const string TypePermissionResponse = "permission_response";
```

### 任务 2：sidecar.py 注册 PreToolUse hook（30 min）

**文件**：[sidecar.py](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.Sidecar/sidecar.py)

**2.1** 定义高风险工具集合
```python
HIGH_RISK_TOOLS = {"execute_vba", "execute_python", "rollback", "clean_data", "remove_duplicates"}
```

**2.2** 定义权限请求的异步等待机制（用 asyncio.Future + 字典）
```python
_pending_permissions = {}  # request_id -> asyncio.Future

async def _request_permission(tool_name: str, args: dict) -> str:
    """向 C# 请求权限，返回 'allow' 或 'deny'"""
    request_id = str(uuid.uuid4())
    future = asyncio.get_event_loop().create_future()
    _pending_permissions[request_id] = future
    # 发送到 C#
    await write_message({
        "type": "permission_request",
        "tool": tool_name,
        "args": args,
        "request_id": request_id,
    })
    # 等待 C# 回复（超时 5 分钟，用户可能走开）
    try:
        decision = await asyncio.wait_for(future, timeout=300)
    except asyncio.TimeoutError:
        decision = "deny"
    finally:
        _pending_permissions.pop(request_id, None)
    return decision
```

**2.3** 定义 PreToolUse hook 回调
```python
async def pre_tool_use_hook(input_data: dict) -> dict:
    tool_name = input_data.get("tool_name", "")
    # 只拦截 mcp__excel__ 前缀的高风险工具
    if not tool_name.startswith("mcp__excel__"):
        return {"continue_": True}
    bare_name = tool_name.replace("mcp__excel__", "")
    if bare_name not in HIGH_RISK_TOOLS:
        return {"continue_": True}  # 低风险直接放行
    # 高风险：请求用户确认
    tool_input = input_data.get("tool_input", {})
    decision = await _request_permission(bare_name, tool_input)
    return {
        "hookSpecificOutput": {
            "hookEventName": "PreToolUse",
            "permissionDecision": decision,  # "allow" 或 "deny"
            "permissionDecisionReason": "用户已确认" if decision == "allow" else "用户拒绝执行",
        }
    }
```

**2.4** 在 ClaudeAgentOptions 注册 hook
```python
from claude_agent_sdk.types import HookMatcher

options = ClaudeAgentOptions(
    # ...现有参数...
    hooks={
        "PreToolUse": [HookMatcher(matcher=None, hooks=[pre_tool_use_hook])],
    },
)
```

**2.5** 处理 C# 回复的 permission_response 消息
```python
# 在 reader 协程的消息处理中
if msg_type == "permission_response":
    request_id = msg.get("request_id")
    decision = msg.get("decision", "deny")
    future = _pending_permissions.get(request_id)
    if future and not future.done():
        future.set_result(decision)
```

### 任务 3：PythonSidecar.cs 处理 permission_request（20 min）

**文件**：[PythonSidecar.cs](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Sidecar/PythonSidecar.cs)

**3.1** 在 OnStdoutLine 的 switch 中新增 case
```csharp
case SidecarProtocol.TypePermissionRequest:
    var reqId = root.GetProperty("request_id").GetString();
    var tool = root.GetProperty("tool").GetString();
    var permArgs = root.GetProperty("args");
    SafeBeginInvoke(() => OnPermissionRequest?.Invoke(this, reqId, tool, permArgs));
    break;
```

**3.2** 新增事件和回复方法
```csharp
public event Action<PythonSidecar, string, string, JsonElement> OnPermissionRequest;

public void SendPermissionResponse(string requestId, string decision)
{
    var msg = new { type = SidecarProtocol.TypePermissionResponse, request_id = requestId, decision };
    WriteLine(JsonSerializer.Serialize(msg, _jsonOptions));
}
```

### 任务 4：MessageBridge.cs 转发到 WebView（10 min）

**文件**：[MessageBridge.cs](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Bridge/MessageBridge.cs)

**4.1** 订阅 OnPermissionRequest 事件，转发到 WebView
```csharp
sidecar.OnPermissionRequest += (s, reqId, tool, args) =>
{
    SendToUi(new { type = "permission_request", payload = new {
        request_id = reqId,
        tool,
        args = args.Deserialize<object>(),
    }});
};
```

**4.2** 处理 WebView 发回的 permission_response
```csharp
// 在 HandleWebViewMessage 中
case "permission_response":
    var respPayload = JsonDocument.Parse(message).RootElement.GetProperty("payload");
    var respReqId = respPayload.GetProperty("request_id").GetString();
    var decision = respPayload.GetProperty("decision").GetString();
    activeSession?.Sidecar?.SendPermissionResponse(respReqId, decision);
    break;
```

### 任务 5：前端 PermissionDrawer 组件（40 min）

**文件**：新建 [PermissionDrawer.tsx](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.UI/src/components/PermissionDrawer.tsx)

**5.1** 组件结构
```tsx
interface PermissionDrawerProps {
  visible: boolean
  tool: string
  args: Record<string, any>
  onAllow: () => void
  onDeny: () => void
}

export function PermissionDrawer({ visible, tool, args, onAllow, onDeny }: PermissionDrawerProps) {
  if (!visible) return null
  const desc = TOOL_DESC[tool] || `AI 想执行 ${tool}`
  return (
    <div className="permission-drawer">
      <div className="permission-drawer-header">
        <span className="permission-icon">⚠</span>
        <span>{desc}</span>
      </div>
      {args && Object.keys(args).length > 0 && (
        <div className="permission-args">
          {Object.entries(args).slice(0, 5).map(([k, v]) => (
            <div key={k} className="permission-arg">
              <span className="arg-key">{k}:</span>
              <span className="arg-val">{String(v).slice(0, 80)}</span>
            </div>
          ))}
        </div>
      )}
      <div className="permission-actions">
        <button className="perm-deny-btn" onClick={onDeny}>拒绝</button>
        <button className="perm-allow-btn" onClick={onAllow}>允许此次</button>
      </div>
    </div>
  )
}
```

**5.2** CSS：从输入框上方滑出（slide-up 动画）
```css
.permission-drawer {
  max-height: 0;
  overflow: hidden;
  transition: max-height 0.3s ease-out;
  background: var(--color-bg);
  border-top: 1px solid var(--color-border);
}
.permission-drawer.visible {
  max-height: 300px;  /* 足够显示内容 */
}
/* 内容样式 */
.permission-drawer-header { ... }
.permission-args { ... }
.permission-actions { display: flex; gap: 8px; }
.perm-allow-btn { background: var(--color-primary); color: white; }
.perm-deny-btn { background: var(--color-bg-secondary); }
```

### 任务 6：App.tsx 集成 PermissionDrawer（15 min）

**文件**：[App.tsx](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.UI/src/App.tsx)

**6.1** 新增状态
```tsx
const [permission, setPermission] = useState<{
  visible: boolean
  requestId?: string
  tool?: string
  args?: Record<string, any>
}>({ visible: false })
```

**6.2** 在 onHostMessage 中处理 permission_request
```tsx
else if (data.type === 'permission_request') {
  setPermission({
    visible: true,
    requestId: data.payload.request_id,
    tool: data.payload.tool,
    args: data.payload.args,
  })
}
```

**6.3** 渲染位置：InputArea **上方**（MessageList 和 InputArea 之间）
```tsx
<MessageList ... />
<PermissionDrawer
  visible={permission.visible}
  tool={permission.tool || ''}
  args={permission.args || {}}
  onAllow={() => {
    sendToHost({ type: 'permission_response', payload: {
      request_id: permission.requestId, decision: 'allow'
    }})
    setPermission({ visible: false })
  }}
  onDeny={() => {
    sendToHost({ type: 'permission_response', payload: {
      request_id: permission.requestId, decision: 'deny'
    }})
    setPermission({ visible: false })
  }}
/>
<InputArea ... />
```

### 任务 7：移除 ToolDispatcher 中的 MessageBox 确认（10 min）

**文件**：[ToolDispatcher.cs](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Sidecar/ToolDispatcher.cs)

**7.1** 移除 PromptUserApproval 方法（或保留但不调用）

**7.2** 移除 Execute 方法中的 SecurityGateway 检查
```csharp
// 删除这段：
if (_securityGateway != null && _securityGateway.RequiresVerification(toolName))
{
    bool approved = PromptUserApproval(toolName, args);
    if (!approved) { ... }
}
```

权限确认改由 PreToolUse hook 在 SDK 层处理，ToolDispatcher 只负责执行。

### 任务 8：移除 ExecutionGuard（10 min）

**文件**：[ToolDispatcher.cs](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/Sidecar/ToolDispatcher.cs)

移除 `ExecutionGuardActive` 相关代码（不再需要，因为 UI 线程不阻塞，不会触发误关闭）：
- 删除 `ExecutionGuardActive` 静态字段
- 删除 Execute 方法中的 `previousGuard` / `ExecutionGuardActive = true/false` 逻辑

**文件**：[ThisAddIn.cs](file:///c:/Users/qinju/Desktop/AIProject/DeepExcel/src/DeepExcel.AddIn/ThisAddIn.cs)

移除 OnWorkbookBeforeClose 中的 ExecutionGuard 检查：
```csharp
// 删除这段：
if (Sidecar.ToolDispatcher.ExecutionGuardActive)
{
    Log("WorkbookBeforeClose SKIPPED ...");
    return;
}
```

## 五、验证清单

- [ ] Python 语法检查通过
- [ ] C# 编译成功
- [ ] 前端构建成功
- [ ] 简单操作（写公式、排序）不弹抽屉
- [ ] VBA 执行时抽屉从输入框上方滑出
- [ ] 点"允许"后 VBA 执行成功
- [ ] 点"拒绝"后 AI 收到拒绝消息
- [ ] 确认期间 Excel 仍可操作（不卡死）
- [ ] 确认期间面板不消失
- [ ] 连续多次 VBA 确认不崩溃

## 六、风险

1. **PreToolUse hook 异步性**：hook 是异步的，但 SDK 是否支持长时间等待（5 分钟超时）需验证
2. **hook 只对 MCP 工具触发**：需确认 `matcher=None` 是否匹配 `mcp__excel__*` 工具
3. **移除 ExecutionGuard 后**：如果 VBA 本身触发 WorkbookBeforeClose（非误触发），需要确保正常清理逻辑健壮

## 七、工作量估算

| 任务 | 文件 | 时间 |
|---|---|---|
| 任务 1 | SidecarProtocol.cs | 5 min |
| 任务 2 | sidecar.py | 30 min |
| 任务 3 | PythonSidecar.cs | 20 min |
| 任务 4 | MessageBridge.cs | 10 min |
| 任务 5 | PermissionDrawer.tsx + CSS | 40 min |
| 任务 6 | App.tsx | 15 min |
| 任务 7 | ToolDispatcher.cs | 10 min |
| 任务 8 | ToolDispatcher.cs + ThisAddIn.cs | 10 min |
| **合计** | **8 个文件** | **~140 min** |

## 八、待确认

1. 是否按此计划执行？
2. 是否保留"允许并记住"选项（同一工具本次会话内不再询问）？还是每次都问？
