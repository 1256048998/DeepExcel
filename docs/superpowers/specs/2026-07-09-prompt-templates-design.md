# 常用提示词（Slash Commands）功能设计

> **日期**：2026-07-09
> **状态**：已批准，进入实现

## 1. 需求

用户希望能将常用需求保存为"提示词模板"，下次通过 `/` 快捷调用，参考 Claude Code 的斜杠命令交互。

**来源**：
- 从历史对话保存（用户发的某条消息 → 保存为模板）
- 手动创建（在管理面板输入标题+内容）

**触发**：输入框输入 `/` 弹出下拉列表，方向键选择 + Enter 确认，ESC 关闭。

## 2. 设计

### 2.1 数据模型

```ts
interface PromptTemplate {
  id: string          // UUID
  title: string       // 显示名称（如 "清洗数据"）
  content: string     // 提示词内容（如 "清洗A列数据，去除空格和特殊字符"）
  createdAt: number   // 创建时间戳
  source: 'history' | 'manual'  // 来源
}
```

存储：`localStorage` key = `deepexcel_prompts`，值为 `PromptTemplate[]`。
理由：纯前端功能，无需 C# 端持久化；localStorage 简单可靠，跨会话保留。

### 2.2 组件

**PromptDropdown.tsx**（新增）— `/` 触发的下拉列表
- 输入框值为 `/` 或 `/xxx` 时显示
- 模糊匹配 title 和 content
- 键盘导航：↑↓ 选择，Enter 确认，ESC 关闭
- 选中后：用 content 填充输入框（不自动发送，用户可编辑后发送）
- 顶部固定项：`+ 新建提示词`（跳转到管理面板）

**PromptManager.tsx**（新增）— 管理面板（模态框）
- 列表展示所有提示词（title + content 预览 + 来源标签）
- 新增：标题 + 内容输入框
- 编辑：点击列表项进入编辑
- 删除：每项右侧删除按钮
- 从历史保存时：预填充 content，用户编辑标题后保存

**MessageItem 悬停按钮**（修改 MessageList.tsx）
- 用户消息悬停时显示"保存为提示词"小按钮（回形针图标旁）
- 点击后打开 PromptManager，content 预填该消息内容

### 2.3 交互流程

```
输入 `/` → 显示 PromptDropdown
  ├─ 输入 `/清洗` → 模糊匹配，过滤显示
  ├─ ↑↓ 选择 → 高亮
  ├─ Enter → content 填入输入框，关闭下拉
  ├─ ESC → 关闭下拉
  └─ 选择"+ 新建" → 打开 PromptManager

历史消息悬停 → 显示"保存为提示词"按钮
  └─ 点击 → 打开 PromptManager（content 预填）

PromptManager
  ├─ 列表 + 增删改
  └─ 保存 → 更新 localStorage + 关闭面板
```

### 2.4 文件清单

| 文件 | 操作 | 职责 |
|------|------|------|
| `src/components/PromptDropdown.tsx` | 新增 | `/` 触发的下拉列表 |
| `src/components/PromptManager.tsx` | 新增 | 管理面板（增删改） |
| `src/utils/prompts.ts` | 新增 | localStorage 读写 + 类型定义 |
| `src/components/InputArea.tsx` | 修改 | 集成 PromptDropdown |
| `src/components/MessageList.tsx` | 修改 | 用户消息悬停"保存为提示词"按钮 |
| `src/App.tsx` | 修改 | PromptManager 状态管理 |
| `src/styles.css` | 修改 | 下拉列表 + 管理面板样式 |

### 2.5 样式设计

- **PromptDropdown**：绝对定位浮在输入框上方，浅灰背景，圆角 6px，最大高度 240px 滚动
- **选中项**：浅蓝背景 + 主色文字
- **PromptManager**：居中模态框（复用现有 conv-panel 样式风格）
- **悬停按钮**：用户消息气泡右下角，小图标，半透明，hover 时显现

### 2.6 非目标

- 不做云端同步（纯本地）
- 不做分类/标签（YAGNI，列表够用）
- 不做导入导出（YAGNI）
- 不自动发送（用户需确认内容后手动 Enter）

## 3. 验证

- 输入 `/` 弹出列表，含"+ 新建"
- 创建提示词后，`/` 列表能匹配到
- 历史消息悬停显示保存按钮，保存后出现在列表
- 删除提示词后列表更新
- ESC 关闭下拉，Enter 填充输入框
