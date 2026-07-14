# 会话内自动压缩 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 当 CLI autocompact 触发时，在对话中插入提示卡，让用户知道发生了压缩。

**Architecture:** 依赖 Claude Agent SDK 原生 autocompact（CLI 内置）。sidecar 在每轮对话后调 `get_context_usage()`，比较前后 percentage 检测压缩发生（percentage 突然下降 >40%），发 compacted 消息给前端，前端插入提示卡。

**Tech Stack:** Python（sidecar.py）、TypeScript/React（前端）、Claude Agent SDK v0.2.109+

## Global Constraints

- 不自建压缩逻辑，依赖 CLI 原生 autocompact
- 不做 token 进度条（用户确认：Excel 用户不太会超上下文，常驻进度条是噪音）
- 不自定义压缩阈值，用 CLI 默认值
- 压缩提示卡用浅灰背景（遵循 UI 规范，不用 AI 风格蓝紫色）
- SDK 无 isCompactSummary 标记，需通过 percentage 下降检测压缩
- sidecar.py 改动需保持 cp1252 兼容（英文系统）：错误消息不含中文/emoji

---

### Task 1: sidecar.py — 压缩检测 + compacted 消息

**Files:**
- Modify: `src/DeepExcel.Sidecar/sidecar.py`（handle_sdk_message 函数 + run_agent_loop 函数）

**Interfaces:**
- Consumes: `ClaudeSDKClient.get_context_usage()` — 返回 `{ totalTokens, maxTokens, percentage, ... }`
- Produces: IPC 消息 `{"type": "compacted", "prev_percentage": float, "curr_percentage": float}`

**关键设计决策：**
- SDK 源码中无 `isCompactSummary` 标记（已验证），改用 percentage 下降检测
- 阈值：curr_percentage < prev_percentage * 0.6（下降超 40%）判定为压缩发生
- 首轮无 prev_percentage，不检测

- [ ] **Step 1: 在 handle_sdk_message 中添加压缩检测逻辑**

在 `handle_sdk_message` 函数的 `ResultMessage` 分支末尾（发送 stream_end 之后），添加 `get_context_usage()` 调用和压缩检测。

在文件顶部添加全局变量：

```python
# 压缩检测：记录上一轮 context usage percentage
_prev_context_percentage: float | None = None
```

修改 `handle_sdk_message` 的 `ResultMessage` 分支，在 `await write_message({"type": "stream_end", ...})` 之后添加：

```python
    elif isinstance(response, ResultMessage):
        usage = getattr(response, "usage", None)
        if isinstance(usage, dict):
            in_tok = usage.get("input_tokens", 0)
            out_tok = usage.get("output_tokens", 0)
        elif usage:
            in_tok = getattr(usage, "input_tokens", 0)
            out_tok = getattr(usage, "output_tokens", 0)
        else:
            in_tok = 0
            out_tok = 0
        await write_message({
            "type": "stream_end",
            "input_tokens": in_tok,
            "output_tokens": out_tok,
        })
        # ★ 压缩检测：每轮结束后查 context usage，percentage 突然下降说明 autocompact 发生了
        global _prev_context_percentage
        try:
            context_usage = await client.get_context_usage()
            curr_pct = context_usage.get("percentage", 0) if isinstance(context_usage, dict) else 0
            if _prev_context_percentage is not None and curr_pct < _prev_context_percentage * 0.6:
                # percentage 下降超 40%，判定为压缩发生
                await write_message({
                    "type": "compacted",
                    "prev_percentage": _prev_context_percentage,
                    "curr_percentage": curr_pct,
                })
            _prev_context_percentage = curr_pct
        except Exception as ctx_e:
            sys.stderr.write(f"[sidecar] get_context_usage failed: {type(ctx_e).__name__}: {ctx_e}\n")
            sys.stderr.flush()
```

注意：`handle_sdk_message` 需要访问 `client` 变量。当前函数签名是 `handle_sdk_message(response)`，需要改为 `handle_sdk_message(response, client)`。

- [ ] **Step 2: 修改 handle_sdk_message 签名，传入 client**

将函数签名从：
```python
async def handle_sdk_message(response):
```
改为：
```python
async def handle_sdk_message(response, client):
```

在 `run_agent_loop` 中调用处（约 L718）：
```python
async for response in client.receive_response():
    await handle_sdk_message(response)
```
改为：
```python
async for response in client.receive_response():
    await handle_sdk_message(response, client)
```

- [ ] **Step 3: 重置 _prev_context_percentage**

在 `run_agent_loop` 的重置区域（约 L621-L624，重置 `_had_partial_text` 等处），添加：

```python
        global _prev_context_percentage
        _prev_context_percentage = None
```

不，这里不应该重置——_prev_context_percentage 应该跨轮次保留（用来比较前后变化）。只有新建对话时才重置。

在文件顶部全局变量区添加即可，不需要在 run_agent_loop 中重置。跳过此步。

- [ ] **Step 4: 验证 sidecar 语法**

Run: `cd src\DeepExcel.Sidecar && python -c "import ast; ast.parse(open('sidecar.py', encoding='utf-8').read()); print('syntax OK')"`
Expected: `syntax OK`

- [ ] **Step 5: Commit**

```bash
git add src/DeepExcel.Sidecar/sidecar.py
git commit -m "feat: detect autocompact via get_context_usage percentage drop"
```

---

### Task 2: 前端类型定义 — Message 类型加 'compacted'

**Files:**
- Modify: `src/DeepExcel.UI/src/types.ts`

**Interfaces:**
- Produces: `Message.type` 新增 `'compacted'` 值

- [ ] **Step 1: 修改 Message 类型**

将 `types.ts` 第 7 行：
```typescript
  type?: 'clarify'
```
改为：
```typescript
  type?: 'clarify' | 'compacted'
```

- [ ] **Step 2: Commit**

```bash
git add src/DeepExcel.UI/src/types.ts
git commit -m "feat: add 'compacted' message type"
```

---

### Task 3: App.tsx — 处理 compacted 消息

**Files:**
- Modify: `src/DeepExcel.UI/src/App.tsx`（onHostMessage 回调）

**Interfaces:**
- Consumes: IPC 消息 `{"type": "compacted", "prev_percentage": float, "curr_percentage": float}`
- Produces: 插入一条 `type: 'compacted'` 的 Message 到 messages 数组

- [ ] **Step 1: 在 onHostMessage 中添加 compacted 处理**

在 `App.tsx` 的 `onHostMessage` 回调中，找到 `} else if (data.type === 'error') {` 之前（约 L186），添加：

```typescript
      } else if (data.type === 'compacted') {
        // ★ autocompact 触发：插入压缩提示卡
        setMessages(prev => [...prev, {
          role: 'assistant',
          content: '对话已自动压缩，保留了关键上下文。',
          type: 'compacted'
        }])
```

- [ ] **Step 2: 验证 TypeScript 编译**

Run: `cd src\DeepExcel.UI && npx tsc --noEmit`
Expected: 0 errors

- [ ] **Step 3: Commit**

```bash
git add src/DeepExcel.UI/src/App.tsx
git commit -m "feat: handle compacted message in App"
```

---

### Task 4: MessageList.tsx — 渲染压缩提示卡

**Files:**
- Modify: `src/DeepExcel.UI/src/components/MessageList.tsx`

**Interfaces:**
- Consumes: `Message.type === 'compacted'`

- [ ] **Step 1: 在消息渲染中添加 compacted 类型处理**

找到 MessageList.tsx 中消息渲染逻辑（通常是根据 `msg.role` 和 `msg.type` 判断渲染方式的区域）。在 `type === 'clarify'` 判断之后，添加 compacted 类型处理：

```tsx
      {msg.type === 'compacted' ? (
        <div className="message compacted-hint">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: 6, verticalAlign: 'middle' }}>
            <path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/>
            <polyline points="21 3 21 8 16 8"/>
          </svg>
          {msg.content}
        </div>
      ) : msg.type === 'clarify' ? (
```

注意：需要确认 MessageList.tsx 现有的条件渲染结构，把 compacted 判断放在最前面（因为它不依赖 role）。

- [ ] **Step 2: 验证 TypeScript 编译**

Run: `cd src\DeepExcel.UI && npx tsc --noEmit`
Expected: 0 errors

- [ ] **Step 3: Commit**

```bash
git add src/DeepExcel.UI/src/components/MessageList.tsx
git commit -m "feat: render compacted hint card"
```

---

### Task 5: styles.css — 压缩提示卡样式

**Files:**
- Modify: `src/DeepExcel.UI/src/styles.css`

- [ ] **Step 1: 添加 .compacted-hint 样式**

在 `styles.css` 末尾添加：

```css
/* === 压缩提示卡（autocompact 触发时显示） === */
.message.compacted-hint {
  margin: 4px auto;
  padding: 6px 12px;
  background: #f3f4f6;
  border: 1px solid #e5e7eb;
  border-radius: 6px;
  font-size: 12px;
  color: #6b7280;
  text-align: center;
  max-width: 90%;
}
```

- [ ] **Step 2: Commit**

```bash
git add src/DeepExcel.UI/src/styles.css
git commit -m "style: add compacted-hint card style"
```

---

### Task 6: 构建验证 + 部署

**Files:**
- 无文件改动，仅构建和部署

- [ ] **Step 1: 前端构建**

Run: `cd src\DeepExcel.UI && npm run build`
Expected: tsc + vite 0 errors

- [ ] **Step 2: 复制 WebViewAssets 到 bin/Release**

Run: `Copy-Item "src\DeepExcel.AddIn\WebViewAssets\assets\*" "src\DeepExcel.AddIn\bin\Release\WebViewAssets\assets\" -Force -Recurse`
Expected: 无错误

- [ ] **Step 3: 复制 sidecar.py 到 bin/Release**

Run: `Copy-Item "src\DeepExcel.Sidecar\sidecar.py" "src\DeepExcel.AddIn\bin\Release\sidecar\sidecar.py" -Force`
Expected: 无错误

- [ ] **Step 4: C# 重新编译（触发 MSBuild 复制资源）**

Run: `& "$env:USERPROFILE\Desktop\AIProject\DeepExcel\scripts\build-csc.bat"`
Expected: Build SUCCESSFUL

- [ ] **Step 5: 注册 CLSID**

Run: `& "$env:USERPROFILE\Desktop\AIProject\DeepExcel\scripts\register-user.ps1"`
Expected: Registration successful

- [ ] **Step 6: 手动验证**

关闭所有 Excel 窗口 → 重新打开 Excel → 点"打开面板" → 进行长对话（超过 20 轮）→ 观察是否出现压缩提示卡。

注意：由于 Excel 用户通常不会超上下文，这个功能很难在日常使用中触发。验证重点是：
1. 短对话不受影响（不出现误报）
2. 代码无语法错误
3. sidecar 启动正常

- [ ] **Step 7: 最终 Commit**

```bash
git add -A
git commit -m "build: context compaction feature complete"
```
