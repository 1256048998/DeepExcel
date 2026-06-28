# DeepExcel P3 商业化实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task.

**Goal:** 实现完整商业化能力，积分制计费、用户体系、遥测

**Architecture:** 
- 在P2基础上增加后端服务
- 实现用户认证和积分账本
- 添加使用遥测和数据分析
- 支持企业License管理

**Tech Stack:** 
- C# / .NET Framework 4.7.2+ / VSTO
- Node.js / Express 或 Python / FastAPI (后端)
- PostgreSQL / MySQL (数据库)
- React + TypeScript (前端)

---

## Global Constraints

- 所有计费必须在后端完成，不能信任客户端
- API Key必须加密存储
- 用户数据必须符合隐私法规
- 遥测数据必须匿名化

---

## Task 1: 后端服务架构

**Files:**
- Create: `src/DeepExcel.Backend/src/index.ts` - Express入口
- Create: `src/DeepExcel.Backend/src/routes/auth.ts` - 认证路由
- Create: `src/DeepExcel.Backend/src/routes/billing.ts` - 计费路由
- Create: `src/DeepExcel.Backend/src/routes/models.ts` - 模型代理路由
- Create: `src/DeepExcel.Backend/src/middleware/auth.ts` - 认证中间件
- Create: `src/DeepExcel.Backend/src/services/creditService.ts` - 积分服务
- Create: `src/DeepExcel.Backend/src/services/modelProxy.ts` - 模型代理服务
- Create: `src/DeepExcel.Backend/src/db/schema.sql` - 数据库Schema
- Create: `src/DeepExcel.Backend/src/config.example.env` - 环境变量示例

**Interfaces:**
- 提供REST API给客户端
- 管理用户账户和积分

- [ ] **Step 1: 创建后端项目结构**

```
src/DeepExcel.Backend/
├─ src/
│  ├─ index.ts
│  ├─ routes/
│  │  ├─ auth.ts
│  │  ├─ billing.ts
│  │  └─ models.ts
│  ├─ middleware/
│  │  └─ auth.ts
│  ├─ services/
│  │  ├─ creditService.ts
│  │  └─ modelProxy.ts
│  ├─ db/
│  │  └─ schema.sql
│  └─ config.example.env
├─ package.json
└─ tsconfig.json
```

- [ ] **Step 2: 创建数据库Schema**

```sql
-- users表
CREATE TABLE users (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    email VARCHAR(255) UNIQUE NOT NULL,
    password_hash VARCHAR(255) NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- credits账本表
CREATE TABLE credits (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    user_id UUID REFERENCES users(id),
    balance INT NOT NULL DEFAULT 0,
    total_purchased INT NOT NULL DEFAULT 0,
    total_consumed INT NOT NULL DEFAULT 0,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- credit_transactions积分变动记录
CREATE TABLE credit_transactions (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    user_id UUID REFERENCES users(id),
    amount INT NOT NULL,  -- 正数=充值，负数=消费
    balance_after INT NOT NULL,
    reason VARCHAR(255),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- api_keys用户API Key表
CREATE TABLE api_keys (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    user_id UUID REFERENCES users(id),
    key_hash VARCHAR(255) NOT NULL,
    label VARCHAR(100),
    last_used_at TIMESTAMP,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- usage_log使用日志
CREATE TABLE usage_log (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    user_id UUID REFERENCES users(id),
    model VARCHAR(50) NOT NULL,
    input_tokens INT NOT NULL,
    output_tokens INT NOT NULL,
    cost DECIMAL(10,4) NOT NULL,
    endpoint VARCHAR(100),
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- subscriptions订阅表
CREATE TABLE subscriptions (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    user_id UUID REFERENCES users(id),
    plan VARCHAR(50) NOT NULL,  -- starter, pro, premium
    status VARCHAR(50) NOT NULL,  -- active, cancelled, expired
    current_period_start TIMESTAMP,
    current_period_end TIMESTAMP,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);
```

- [ ] **Step 3: 创建后端入口**

```typescript
// index.ts
import express from 'express';
import cors from 'cors';
import authRoutes from './routes/auth';
import billingRoutes from './routes/billing';
import modelRoutes from './routes/models';

const app = express();

app.use(cors());
app.use(express.json());

// 路由
app.use('/api/auth', authRoutes);
app.use('/api/billing', billingRoutes);
app.use('/api/models', modelRoutes);

// 健康检查
app.get('/health', (req, res) => {
    res.json({ status: 'ok' });
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
    console.log(`DeepExcel Backend running on port ${PORT}`);
});
```

- [ ] **Step 4: 创建认证路由**

```typescript
// routes/auth.ts
import { Router } from 'express';
import bcrypt from 'bcrypt';
import { db } from '../db';
import { v4 as uuid } from 'uuid';

const router = Router();

// 注册
router.post('/register', async (req, res) => {
    const { email, password } = req.body;
    
    const hashed = await bcrypt.hash(password, 10);
    
    try {
        const userId = uuid();
        await db.query(
            'INSERT INTO users (id, email, password_hash) VALUES ($1, $2, $3)',
            [userId, email, hashed]
        );
        
        // 创建初始积分账本
        await db.query(
            'INSERT INTO credits (user_id, balance) VALUES ($1, 0)',
            [userId]
        );
        
        res.json({ success: true, userId });
    } catch (err) {
        res.status(400).json({ error: 'Email already exists' });
    }
});

// 登录
router.post('/login', async (req, res) => {
    const { email, password } = req.body;
    
    const result = await db.query(
        'SELECT * FROM users WHERE email = $1',
        [email]
    );
    
    if (result.rows.length === 0) {
        return res.status(401).json({ error: 'Invalid credentials' });
    }
    
    const user = result.rows[0];
    const valid = await bcrypt.compare(password, user.password_hash);
    
    if (!valid) {
        return res.status(401).json({ error: 'Invalid credentials' });
    }
    
    // 生成API Key
    const apiKey = uuid();
    await db.query(
        'INSERT INTO api_keys (user_id, key_hash, label) VALUES ($1, $2, $3)',
        [user.id, apiKey, 'Default']
    );
    
    res.json({ 
        success: true, 
        apiKey,
        userId: user.id 
    });
});

export default router;
```

- [ ] **Step 5: 创建积分服务**

```typescript
// services/creditService.ts
import { db } from '../db';

export class CreditService {
    // 计算消费
    async consume(userId: string, amount: number, reason: string): Promise<boolean> {
        const result = await db.query(
            'SELECT balance FROM credits WHERE user_id = $1 FOR UPDATE',
            [userId]
        );
        
        if (result.rows.length === 0) return false;
        
        const currentBalance = result.rows[0].balance;
        if (currentBalance < amount) return false; // 积分不足
        
        const newBalance = currentBalance - amount;
        
        await db.query('BEGIN');
        try {
            await db.query(
                'UPDATE credits SET balance = $1, total_consumed = total_consumed + $2 WHERE user_id = $3',
                [newBalance, amount, userId]
            );
            await db.query(
                'INSERT INTO credit_transactions (user_id, amount, balance_after, reason) VALUES ($1, $2, $3, $4)',
                [userId, -amount, newBalance, reason]
            );
            await db.query('COMMIT');
            return true;
        } catch (err) {
            await db.query('ROLLBACK');
            return false;
        }
    }
    
    // 充值
    async recharge(userId: string, amount: number): Promise<boolean> {
        await db.query('BEGIN');
        try {
            await db.query(
                'UPDATE credits SET balance = balance + $1, total_purchased = total_purchased + $1 WHERE user_id = $2',
                [amount, userId]
            );
            const result = await db.query('SELECT balance FROM credits WHERE user_id = $1', [userId]);
            await db.query(
                'INSERT INTO credit_transactions (user_id, amount, balance_after, reason) VALUES ($1, $2, $3, $4)',
                [userId, amount, result.rows[0].balance, 'Purchase']
            );
            await db.query('COMMIT');
            return true;
        } catch (err) {
            await db.query('ROLLBACK');
            return false;
        }
    }
    
    // 获取余额
    async getBalance(userId: string): Promise<number> {
        const result = await db.query(
            'SELECT balance FROM credits WHERE user_id = $1',
            [userId]
        );
        return result.rows[0]?.balance || 0;
    }
}
```

- [ ] **Step 6: 提交**

---

## Task 2: 模型代理服务

**Files:**
- Modify: `src/DeepExcel.Backend/src/services/modelProxy.ts`
- Create: `src/DeepExcel.Backend/src/services/anthropicProxy.ts`
- Create: `src/DeepExcel.Backend/src/services/deepseekProxy.ts`

**Interfaces:**
- 统一封装不同模型API
- 处理计费和遥测

- [ ] **Step 1: 创建模型代理基类**

```typescript
// services/modelProxy.ts
export interface ModelRequest {
    model: string;
    messages: Array<{ role: string; content: string }>;
    tools?: any[];
    max_tokens?: number;
    temperature?: number;
}

export interface ModelResponse {
    content: string;
    usage: {
        input_tokens: number;
        output_tokens: number;
    };
    cost: number; // 本次调用的积分成本
}

export abstract class BaseModelProxy {
    protected abstract endpoint: string;
    protected abstract model: string;
    
    abstract send request(req: ModelRequest): Promise<ModelResponse>;
    
    protected calculateCost(inputTokens: number, outputTokens: number): number {
        // Claude Sonnet: $0.003/1K输入, $0.015/1K输出
        return (inputTokens / 1000) * 0.003 + (outputTokens / 1000) * 0.015;
    }
}
```

- [ ] **Step 2: 创建Claude代理**

```typescript
// services/anthropicProxy.ts
import BaseModelProxy, { ModelRequest, ModelResponse } from './modelProxy';

export class AnthropicProxy extends BaseModelProxy {
    protected endpoint = 'https://api.anthropic.com/v1/messages';
    protected model = 'claude-3-5-sonnet-20241022';
    private apiKey: string;
    
    constructor(apiKey: string) {
        super();
        this.apiKey = apiKey;
    }
    
    async sendRequest(req: ModelRequest): Promise<ModelResponse> {
        const response = await fetch(this.endpoint, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'x-api-key': this.apiKey,
                'anthropic-version': '2023-06-01'
            },
            body: JSON.stringify({
                model: this.model,
                max_tokens: req.max_tokens || 4096,
                messages: req.messages,
                tools: req.tools
            })
        });
        
        const data = await response.json();
        
        return {
            content: data.content[0].text,
            usage: {
                input_tokens: data.usage.input_tokens,
                output_tokens: data.usage.output_tokens
            },
            cost: this.calculateCost(data.usage.input_tokens, data.usage.output_tokens)
        };
    }
}
```

- [ ] **Step 3: 创建DeepSeek代理**

```typescript
// services/deepseekProxy.ts
import BaseModelProxy, { ModelRequest, ModelResponse } from './modelProxy';

export class DeepSeekProxy extends BaseModelProxy {
    protected endpoint = 'https://api.deepseek.com/v1/chat/completions';
    protected model = 'deepseek-chat';
    private apiKey: string;
    
    constructor(apiKey: string) {
        super();
        this.apiKey = apiKey;
    }
    
    async sendRequest(req: ModelRequest): Promise<ModelResponse> {
        const response = await fetch(this.endpoint, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${this.apiKey}`
            },
            body: JSON.stringify({
                model: this.model,
                messages: req.messages,
                max_tokens: req.max_tokens || 4096,
                temperature: req.temperature || 0.7
            })
        });
        
        const data = await response.json();
        
        return {
            content: data.choices[0].message.content,
            usage: {
                input_tokens: data.usage.prompt_tokens,
                output_tokens: data.usage.completion_tokens
            },
            cost: this.calculateCost(data.usage.prompt_tokens, data.usage.completion_tokens)
        };
    }
    
    protected calculateCost(inputTokens: number, outputTokens: number): number {
        // DeepSeek: $0.0001/1K输入, $0.0003/1K输出
        return (inputTokens / 1000) * 0.0001 + (outputTokens / 1000) * 0.0003;
    }
}
```

- [ ] **Step 4: 提交**

---

## Task 3: 计费路由

**Files:**
- Create: `src/DeepExcel.Backend/src/routes/billing.ts`

**Interfaces:**
- 提供积分查询、充值、消耗接口

- [ ] **Step 1: 创建计费路由**

```typescript
// routes/billing.ts
import { Router } from 'express';
import { CreditService } from '../services/creditService';
import { requireAuth } from '../middleware/auth';

const router = Router();
const creditService = new CreditService();

// 获取余额
router.get('/balance', requireAuth, async (req, res) => {
    const balance = await creditService.getBalance(req.userId);
    res.json({ balance });
});

// 充值
router.post('/recharge', requireAuth, async (req, res) => {
    const { amount, paymentId } = req.body; // amount: 积分数量
    
    // 实际应该调用支付网关验证
    const success = await creditService.recharge(req.userId, amount);
    
    if (success) {
        res.json({ success: true, newBalance: await creditService.getBalance(req.userId) });
    } else {
        res.status(400).json({ error: 'Recharge failed' });
    }
});

// 消费记录
router.get('/transactions', requireAuth, async (req, res) => {
    const result = await db.query(
        'SELECT * FROM credit_transactions WHERE user_id = $1 ORDER BY created_at DESC LIMIT 50',
        [req.userId]
    );
    res.json({ transactions: result.rows });
});

export default router;
```

- [ ] **Step 2: 提交**

---

## Task 4: 模型路由（带计费）

**Files:**
- Modify: `src/DeepExcel.Backend/src/routes/models.ts`

**Interfaces:**
- 接收模型请求，计费后代理到实际模型

- [ ] **Step 1: 创建模型路由**

```typescript
// routes/models.ts
import { Router } from 'express';
import { requireAuth } from '../middleware/auth';
import { CreditService } from '../services/creditService';
import { AnthropicProxy } from '../services/anthropicProxy';
import { DeepSeekProxy } from '../services/deepseekProxy';
import { db } from '../db';

const router = Router();
const creditService = new CreditService();

// 统一模型调用接口
router.post('/chat', requireAuth, async (req, res) => {
    const { model, messages, tools, max_tokens, temperature } = req.body;
    
    // 创建对应模型的代理
    let proxy;
    switch (model) {
        case 'claude':
            const claudeKey = await getUserModelKey(req.userId, 'claude');
            proxy = new AnthropicProxy(claudeKey);
            break;
        case 'deepseek':
            const deepseekKey = await getUserModelKey(req.userId, 'deepseek');
            proxy = new DeepSeekProxy(deepseekKey);
            break;
        default:
            return res.status(400).json({ error: 'Unknown model' });
    }
    
    // 调用模型
    const result = await proxy.sendRequest({ model, messages, tools, max_tokens, temperature });
    
    // 扣积分
    const deducted = await creditService.consume(
        req.userId, 
        Math.ceil(result.cost * 100), // 转换为积分（1积分=$0.01）
        `Model: ${model}`
    );
    
    if (!deducted) {
        return res.status(402).json({ error: 'Insufficient credits' });
    }
    
    // 记录使用日志
    await db.query(
        'INSERT INTO usage_log (user_id, model, input_tokens, output_tokens, cost) VALUES ($1, $2, $3, $4, $5)',
        [req.userId, model, result.usage.input_tokens, result.usage.output_tokens, result.cost]
    );
    
    res.json({
        content: result.content,
        usage: result.usage
    });
});

// 获取用户指定模型的API Key
async function getUserModelKey(userId: string, model: string): Promise<string> {
    // 实际从数据库或密钥管理服务获取
    // 这里简化处理
    const result = await db.query(
        'SELECT key_hash FROM api_keys WHERE user_id = $1 AND label = $2',
        [userId, model]
    );
    return result.rows[0]?.key_hash || process.env.DEFAULT_MODEL_KEY;
}

export default router;
```

- [ ] **Step 2: 提交**

---

## Task 5: 客户端集成后端

**Files:**
- Modify: `src/DeepExcel.AddIn/Config/SettingsManager.cs` - 支持后端配置
- Modify: `src/DeepExcel.Models/ClaudeAdapter.cs` - 改为调用后端
- Modify: `src/DeepExcel.Models/DeepSeekAdapter.cs` - 改为调用后端

**Interfaces:**
- 客户端通过后端调用模型
- 支持本地API Key或后端代理

- [ ] **Step 1: 更新设置管理器**

```csharp
public class AppSettings
{
    public ModelSettings Model { get; set; } = new();
    // ...
    
    public bool UseLocalKey { get; set; } = true;  // true=本地直连, false=后端代理
    public string BackendUrl { get; set; } = "http://localhost:3001";
}

public class ModelSettings
{
    public string ActiveModel { get; set; } = "claude";
    public string LocalApiKey { get; set; }  // 本地API Key
}
```

- [ ] **Step 2: 更新ClaudeAdapter使用后端**

```csharp
public class ClaudeAdapter : IModelAdapter
{
    private readonly SettingsManager _settings;
    
    public ClaudeAdapter(SettingsManager settings)
    {
        _settings = settings;
    }
    
    public async Task<ModelResponse> SendMessageAsync(ModelRequest request)
    {
        if (_settings.Settings.UseLocalKey)
        {
            return await SendDirectAsync(request);
        }
        else
        {
            return await SendViaBackendAsync(request);
        }
    }
    
    private async Task<ModelResponse> SendViaBackendAsync(ModelRequest request)
    {
        var client = new HttpClient { BaseAddress = new Uri(_settings.Settings.BackendUrl) };
        
        var payload = new
        {
            model = "claude",
            messages = request.Messages,
            tools = request.AvailableTools,
            max_tokens = 1024
        };
        
        var response = await client.PostAsJsonAsync("/api/models/chat", payload);
        var result = await response.Content.ReadFromJsonAsync<BackendResponse>();
        
        return new ModelResponse { Content = result.Content };
    }
}
```

- [ ] **Step 3: 提交**

---

## Task 6: 遥测与数据分析

**Files:**
- Create: `src/DeepExcel.Backend/src/services/telemetry.ts`
- Create: `src/DeepExcel.Backend/src/routes/analytics.ts`
- Create: `src/DeepExcel.AddIn/Telemetry/TelemetryService.cs`

**Interfaces:**
- 收集匿名使用数据
- 提供使用分析API

- [ ] **Step 1: 创建遥测服务（后端）**

```typescript
// services/telemetry.ts
import { db } from '../db';

export class TelemetryService {
    // 记录事件（匿名化）
    async trackEvent(userId: string, event: string, properties: Record<string, any>) {
        // 匿名化处理
        const anonymousId = this.hashUserId(userId);
        
        await db.query(
            'INSERT INTO telemetry_events (anonymous_id, event, properties, created_at) VALUES ($1, $2, $3, NOW())',
            [anonymousId, event, JSON.stringify(properties)]
        );
    }
    
    // 获取使用统计
    async getUsageStats(userId: string, period: 'day' | 'week' | 'month') {
        const interval = period === 'day' ? '1 day' : period === 'week' ? '7 days' : '30 days';
        
        const result = await db.query(`
            SELECT 
                COUNT(*) as total_requests,
                SUM(input_tokens) as total_input_tokens,
                SUM(output_tokens) as total_output_tokens,
                SUM(cost) as total_cost,
                model
            FROM usage_log 
            WHERE user_id = $1 AND created_at > NOW() - INTERVAL '${interval}'
            GROUP BY model
        `, [userId]);
        
        return result.rows;
    }
    
    private hashUserId(userId: string): string {
        // 简单的匿名化（实际应用更复杂的哈希）
        const crypto = require('crypto');
        return crypto.createHash('sha256').update(userId + 'salt').digest('hex').substring(0, 16);
    }
}
```

- [ ] **Step 2: 创建遥测客户端**

```csharp
// Telemetry/TelemetryService.cs
public class TelemetryService
{
    private readonly string _backendUrl;
    
    public TelemetryService(string backendUrl)
    {
        _backendUrl = backendUrl;
    }
    
    public async Task TrackAsync(string eventName, Dictionary<string, string> properties)
    {
        try
        {
            var client = new HttpClient();
            await client.PostAsJsonAsync($"{_backendUrl}/api/telemetry", new
            {
                event = eventName,
                properties = properties,
                timestamp = DateTime.UtcNow
            });
        }
        catch
        {
            // 遥测失败不影响主流程
        }
    }
}
```

- [ ] **Step 3: 提交**

---

## Task 7: 企业License支持

**Files:**
- Create: `src/DeepExcel.Backend/src/services/licenseService.ts`
- Create: `src/DeepExcel.Backend/src/routes/enterprise.ts`

**Interfaces:**
- 支持企业批量License
- 管理员控制台

- [ ] **Step 1: 创建License服务**

```typescript
// services/licenseService.ts
export class LicenseService {
    // 验证企业License
    async validateLicense(licenseKey: string): Promise<LicenseInfo | null> {
        const result = await db.query(
            'SELECT * FROM enterprise_licenses WHERE license_key = $1 AND status = $2',
            [licenseKey, 'active']
        );
        
        if (result.rows.length === 0) return null;
        
        const license = result.rows[0];
        
        // 检查过期
        if (license.expires_at && new Date(license.expires_at) < new Date()) {
            return null;
        }
        
        return {
            licenseKey,
            companyName: license.company_name,
            maxSeats: license.max_seats,
            expiresAt: license.expires_at
        };
    }
    
    // 激活Seat
    async activateSeat(licenseKey: string, userId: string): Promise<boolean> {
        const license = await this.validateLicense(licenseKey);
        if (!license) return false;
        
        const seatCount = await db.query(
            'SELECT COUNT(*) FROM license_seats WHERE license_key = $1',
            [licenseKey]
        );
        
        if (seatCount.rows[0].count >= license.maxSeats) {
            return false; // 达到最大Seat数
        }
        
        await db.query(
            'INSERT INTO license_seats (license_key, user_id, activated_at) VALUES ($1, $2, NOW())',
            [licenseKey, userId]
        );
        
        return true;
    }
}
```

- [ ] **Step 2: 提交**

---

## Task 8: 安装包制作

**Files:**
- Create: `installer/DeepExcel.iss` - Inno Setup脚本
- Create: `installer/check_webview2.ps1` - WebView2检测脚本
- Create: `installer/check_dotnet.ps1` - .NET检测脚本

**Interfaces:**
- 生成可分发的安装包
- 自动检测依赖

- [ ] **Step 1: 创建Inno Setup脚本**

```iss
#define MyAppName "DeepExcel"
#define MyAppVersion "0.1.0"
#define MyAppPublisher "DeepExcel"
#define MyAppURL "https://deepexcel.ai"
#define MyAppExeName "DeepExcel.AddIn.dll"

[Setup]
AppId={{B5E9-4C7D-9F2A-1E3C-8B7D6F9A0E1C}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
OutputBaseFilename=DeepExcel-Setup-{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "..\src\DeepExcel.AddIn\bin\Release\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs
Source: "WebView2Loader.dll"; DestDir: "{app}"; Flags: ignoreversion

[Run]
Filename: "{app}\check_dotnet.ps1"; StatusMsg: "Checking .NET Framework..."
Filename: "{app}\check_webview2.ps1"; StatusMsg: "Checking WebView2 Runtime..."

[Code]
function IsDotNetInstalled(): Boolean;
var
  ResultCode: Integer;
begin
  Result := Exec('powershell.exe', '-Command', 'Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" -Name Release -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Release', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;
```

- [ ] **Step 2: 创建WebView2检测脚本**

```powershell
# check_webview2.ps1
$edgePath = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
$userDataFolder = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"

if (!(Test-Path $edgePath)) {
    Write-Host "Microsoft Edge not found. WebView2 requires Edge."
    # 下载WebView2安装包
    $installer = "$env:TEMP\MicrosoftEdgeWebview2Setup.exe"
    Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/p/?LinkId=2124703" -OutFile $installer
    Start-Process -FilePath $installer -Args "/silent /install" -Wait
}
```

- [ ] **Step 3: 提交**

---

## P3验证清单

- [ ] 后端服务启动成功
- [ ] 用户注册/登录功能正常
- [ ] 积分充值和消费正常
- [ ] 通过后端代理调用Claude/DeepSeek
- [ ] 使用日志正确记录
- [ ] 企业License激活和管理
- [ ] 安装包成功生成
- [ ] 首次运行自动安装依赖（.NET/WebView2）
