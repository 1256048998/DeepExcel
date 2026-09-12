// src/DeepExcel.Wps/model-service.js
// 在线拉取厂商模型列表 + 测试 API Key（对应 C# MessageBridge 的
// HandleRefreshModels / HandleTestApiKey / BuildModelEndpointCandidates / ParseModelIds）
//
// 与 C# 端保持同一套规则，避免两个宿主行为不一致：
//   - 厂商已知端点表（对话地址和模型列表地址常常不同路径）
//   - 逐层剥离 /anthropic、/compatible-mode 等兼容层后缀推导候选端点
//   - Anthropic 风格分页（has_more + last_id）
//   - 保持接口返回顺序，不做字母排序（新模型通常在最前）
//   - SSRF 防护：只允许公网 http(s)

const http = require('http')
const https = require('https')
const { URL } = require('url')

const MAX_ENDPOINT_CANDIDATES = 5
const REQUEST_TIMEOUT_MS = 8000
const TEST_TIMEOUT_MS = 15000

/** 各厂商已知的模型列表端点（与 C# KnownModelEndpoints 一致） */
const KNOWN_MODEL_ENDPOINTS = {
  anthropic: ['https://api.anthropic.com/v1/models'],
  deepseek: ['https://api.deepseek.com/models'],
  stepfun: ['https://api.stepfun.com/v1/models'],
  openai: ['https://api.openai.com/v1/models'],
  kimi: ['https://api.moonshot.cn/v1/models'],
  qwen: ['https://dashscope.aliyuncs.com/compatible-mode/v1/models'],
  zhipu: ['https://api.z.ai/api/paas/v4/models', 'https://open.bigmodel.cn/api/paas/v4/models'],
  minimax: ['https://api.minimax.io/v1/models'],
  doubao: ['https://ark.cn-beijing.volces.com/api/v3/models'],
}

/** BaseUrl 中可以安全剥离的兼容层路径段 */
const STRIPPABLE_SEGMENTS = ['anthropic', 'compatible', 'compatible-mode', 'step_plan', 'openai', 'v1']

/** SSRF 防护：拒绝非 http(s)、回环地址与内网网段（与 C# IsValidTestUrl 同口径） */
function isValidTestUrl(rawUrl) {
  let parsed
  try {
    parsed = new URL(rawUrl)
  } catch (e) {
    return false
  }
  if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') return false

  const host = parsed.hostname.toLowerCase()
  if (host === 'localhost' || host.endsWith('.localhost') || host === '::1') return false

  const ipv4 = host.match(/^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/)
  if (ipv4) {
    const [a, b] = [Number(ipv4[1]), Number(ipv4[2])]
    if (a === 127 || a === 10 || a === 0) return false
    if (a === 172 && b >= 16 && b <= 31) return false
    if (a === 192 && b === 168) return false
    if (a === 169 && b === 254) return false  // link-local / 云元数据
  }
  return true
}

function tryGetHost(rawUrl) {
  try {
    return new URL(rawUrl).hostname.toLowerCase()
  } catch (e) {
    return null
  }
}

/** 端点候选：同域已知端点优先 → BaseUrl 推导 → 其余已知端点兜底 */
function buildModelEndpointCandidates(provider, baseUrl) {
  const result = []
  const add = value => {
    if (result.length >= MAX_ENDPOINT_CANDIDATES) return
    if (!value) return
    if (!result.some(item => item.toLowerCase() === value.toLowerCase())) result.push(value)
  }

  const trimmed = String(baseUrl || '').replace(/\/+$/, '')
  const known = KNOWN_MODEL_ENDPOINTS[String(provider || '').toLowerCase()] || null
  const host = tryGetHost(trimmed)

  if (known && host) {
    for (const url of known) {
      if (tryGetHost(url) === host) add(url)
    }
  }

  let current = trimmed
  for (let depth = 0; depth < 3 && current; depth++) {
    if (/\/v1$/i.test(current)) {
      add(current + '/models')
    } else {
      add(current + '/v1/models')
      add(current + '/models')
    }

    const slash = current.lastIndexOf('/')
    if (slash <= current.indexOf('//') + 1) break
    const segment = current.slice(slash + 1)
    if (!STRIPPABLE_SEGMENTS.some(s => s.toLowerCase() === segment.toLowerCase())) break
    current = current.slice(0, slash)
  }

  if (known) for (const url of known) add(url)
  return result
}

/** 解析模型 ID：兼容 OpenAI({data:[{id}]}) / Anthropic / 纯字符串数组，保持原顺序 */
function parseModelIds(body) {
  const models = []
  let doc
  try {
    doc = JSON.parse(body)
  } catch (e) {
    return models
  }

  let array = null
  if (Array.isArray(doc)) array = doc
  else if (doc && Array.isArray(doc.data)) array = doc.data
  else if (doc && Array.isArray(doc.models)) array = doc.models
  if (!array) return models

  for (const item of array) {
    let id = null
    if (typeof item === 'string') id = item
    else if (item && typeof item.id === 'string') id = item.id
    if (!id || !id.trim() || id.length > 200) continue
    if (!models.some(existing => existing.toLowerCase() === id.toLowerCase())) models.push(id)
    if (models.length >= 500) break
  }
  return models
}

/** Anthropic 风格分页游标（has_more + last_id）→ 下一页 URL；无分页返回 null */
function buildNextPageUrl(endpoint, body) {
  try {
    const doc = JSON.parse(body)
    if (!doc || doc.has_more !== true || typeof doc.last_id !== 'string' || !doc.last_id) return null
    const base = endpoint.split('?')[0]
    return `${base}?limit=100&after_id=${encodeURIComponent(doc.last_id)}`
  } catch (e) {
    return null
  }
}

/** 统一的 HTTP 请求封装（Promise） */
function request(url, { method = 'GET', headers = {}, body = null, timeoutMs = REQUEST_TIMEOUT_MS } = {}) {
  return new Promise((resolve, reject) => {
    let parsed
    try {
      parsed = new URL(url)
    } catch (e) {
      reject(new Error('invalid url'))
      return
    }
    const client = parsed.protocol === 'http:' ? http : https
    const req = client.request(url, { method, headers }, res => {
      let data = ''
      res.setEncoding('utf8')
      res.on('data', chunk => { data += chunk })
      res.on('end', () => resolve({ statusCode: res.statusCode, body: data }))
    })
    req.on('error', reject)
    req.setTimeout(timeoutMs, () => {
      req.destroy(new Error('请求超时'))
    })
    if (body) req.write(body)
    req.end()
  })
}

/** 错误体脱敏 + 截断（对应 C# SanitizeErrBody） */
function sanitizeErrBody(body) {
  if (!body) return ''
  let s = String(body)
    .replace(/sk-[A-Za-z0-9_-]{8,}/g, 'sk-***')
    .replace(/Bearer\s+[A-Za-z0-9_\-.]{8,}/g, 'Bearer ***')
  if (s.length > 500) s = s.slice(0, 500) + '...(truncated)'
  return s
}

function authHeaders(apiKey) {
  return {
    Authorization: `Bearer ${apiKey}`,
    'x-api-key': apiKey,
    'anthropic-version': '2023-06-01',
    Accept: 'application/json',
  }
}

/** 请求单个端点并跟随分页；失败返回 { models: null, error } */
async function fetchFromEndpoint(endpoint, apiKey) {
  const all = []
  let url = endpoint
  for (let page = 0; page < 10 && url; page++) {
    let response
    try {
      response = await request(url, { headers: authHeaders(apiKey) })
    } catch (e) {
      return { models: null, error: '模型列表请求失败' }
    }
    if (response.statusCode < 200 || response.statusCode >= 300) {
      return { models: null, error: `HTTP ${response.statusCode}: ${sanitizeErrBody(response.body)}` }
    }
    for (const id of parseModelIds(response.body)) {
      if (!all.some(existing => existing.toLowerCase() === id.toLowerCase())) all.push(id)
    }
    if (all.length >= 500) break
    url = buildNextPageUrl(endpoint, response.body)
  }
  if (all.length === 0) return { models: null, error: '供应商返回成功，但响应中没有模型 ID' }
  return { models: all, error: null }
}

/** 拉取厂商可用模型列表（只读，不落盘；由前端"导入模型"弹窗决定保存什么） */
async function fetchModels({ provider, baseUrl, apiKey }) {
  if (!apiKey) return { success: false, error: '请先输入 API Key' }
  if (!isValidTestUrl(baseUrl)) return { success: false, error: 'Base URL 不安全或格式无效' }

  const candidates = buildModelEndpointCandidates(provider, baseUrl)
  let lastError = null
  for (const endpoint of candidates) {
    if (!isValidTestUrl(endpoint)) continue
    const result = await fetchFromEndpoint(endpoint, apiKey)
    if (result.models && result.models.length > 0) {
      return { success: true, models: result.models, endpoint }
    }
    lastError = result.error || lastError
  }
  return {
    success: false,
    error: lastError || '该供应商暂不支持在线获取模型列表，请继续使用内置模型列表',
  }
}

/** 测试连接：向 Anthropic 兼容的 /v1/messages 发一个 1-token 请求 */
async function testApiKey({ provider, baseUrl, apiKey, model }) {
  if (!apiKey) return { success: false, error: '请先输入 API Key' }
  let url = String(baseUrl || '')
  if (provider === 'deepseek' && !/\/anthropic$/i.test(url)) url = 'https://api.deepseek.com/anthropic'
  if (!isValidTestUrl(url)) {
    return { success: false, error: 'Base URL 不安全：禁止指向内网或非 HTTP(S) 地址' }
  }

  const testUrl = url.replace(/\/+$/, '') + '/v1/messages'
  const payload = JSON.stringify({
    model,
    max_tokens: 1,
    messages: [{ role: 'user', content: 'hi' }],
  })
  const started = Date.now()
  try {
    const response = await request(testUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload),
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
      },
      body: payload,
      timeoutMs: TEST_TIMEOUT_MS,
    })
    const latencyMs = Date.now() - started
    if (response.statusCode >= 200 && response.statusCode < 300) {
      return { success: true, latencyMs, error: null }
    }
    return {
      success: false,
      latencyMs,
      error: `HTTP ${response.statusCode}: ${sanitizeErrBody(response.body)}`,
    }
  } catch (e) {
    return { success: false, error: `连接失败：${e.message}` }
  }
}

module.exports = {
  fetchModels,
  testApiKey,
  buildModelEndpointCandidates,
  parseModelIds,
  buildNextPageUrl,
  isValidTestUrl,
  KNOWN_MODEL_ENDPOINTS,
}
