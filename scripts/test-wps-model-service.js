'use strict'

// model-service.js 单元测试：与 C# 端 ModelRefreshTests 保持同样的断言，
// 确保 Excel / WPS 两个宿主拉取模型列表的行为一致。
//
// 运行：node scripts/test-wps-model-service.js（build-wps.ps1 会自动调用）

const assert = require('assert')
const path = require('path')

const service = require(path.resolve(__dirname, '..', 'src', 'DeepExcel.Wps', 'model-service'))

// ============ 解析：保持厂商返回顺序，不做字母排序 ============
assert.deepStrictEqual(
  service.parseModelIds('{"data":[{"id":"model-b"},{"id":"model-a"}]}'),
  ['model-b', 'model-a'],
)
assert.deepStrictEqual(
  service.parseModelIds('{"models":["alpha","ALPHA","beta"]}'),
  ['alpha', 'beta'],
)
assert.deepStrictEqual(service.parseModelIds('{"result":true}'), [])
assert.deepStrictEqual(service.parseModelIds('not json'), [])

// ============ 分页：Anthropic 的 has_more + last_id ============
assert.strictEqual(
  service.buildNextPageUrl(
    'https://api.anthropic.com/v1/models',
    '{"data":[{"id":"claude-opus-5"}],"has_more":true,"last_id":"claude-opus-5"}',
  ),
  'https://api.anthropic.com/v1/models?limit=100&after_id=claude-opus-5',
)
assert.strictEqual(
  service.buildNextPageUrl('https://api.deepseek.com/models', '{"data":[],"has_more":false}'),
  null,
)
assert.strictEqual(
  service.buildNextPageUrl('https://api.openai.com/v1/models', '{"data":[{"id":"gpt-5.5"}]}'),
  null,
)

// ============ 端点候选：覆盖各厂商真实的模型列表地址 ============
const endpointCases = [
  // 智谱：对话走 /api/anthropic，模型列表在 /api/paas/v4/models
  ['zhipu', 'https://api.z.ai/api/anthropic', 'https://api.z.ai/api/paas/v4/models'],
  // 豆包：对话走 /api/compatible，模型列表在 /api/v3/models
  ['doubao', 'https://ark.cn-beijing.volces.com/api/compatible', 'https://ark.cn-beijing.volces.com/api/v3/models'],
  // 阶跃星辰：对话走 /step_plan，模型列表在 /v1/models
  ['stepfun', 'https://api.stepfun.com/step_plan', 'https://api.stepfun.com/v1/models'],
  // 通义千问：剥离 /anthropic 后的 /compatible-mode/v1/models
  ['qwen', 'https://dashscope.aliyuncs.com/compatible-mode/anthropic', 'https://dashscope.aliyuncs.com/compatible-mode/v1/models'],
  ['deepseek', 'https://api.deepseek.com/anthropic', 'https://api.deepseek.com/models'],
]
for (const [provider, baseUrl, expected] of endpointCases) {
  const candidates = service.buildModelEndpointCandidates(provider, baseUrl)
  assert.ok(
    candidates.includes(expected),
    `${provider}: expected ${expected} in ${JSON.stringify(candidates)}`,
  )
  assert.ok(candidates.length <= 5, `${provider}: too many candidates`)
}

// 自建网关：只按 BaseUrl 推导，且不能越过 host
const gateway = service.buildModelEndpointCandidates('custom', 'https://gateway.example.com/v1')
assert.ok(gateway.includes('https://gateway.example.com/v1/models'))
assert.ok(!gateway.some(u => u.includes('//models')))

// ============ SSRF 防护 ============
assert.strictEqual(service.isValidTestUrl('https://api.openai.com/v1/models'), true)
assert.strictEqual(service.isValidTestUrl('http://localhost:8080/v1/models'), false)
assert.strictEqual(service.isValidTestUrl('http://127.0.0.1/v1/models'), false)
assert.strictEqual(service.isValidTestUrl('http://192.168.1.10/v1/models'), false)
assert.strictEqual(service.isValidTestUrl('http://10.0.0.5/v1/models'), false)
assert.strictEqual(service.isValidTestUrl('http://172.16.3.9/v1/models'), false)
assert.strictEqual(service.isValidTestUrl('http://169.254.169.254/latest/meta-data'), false)
assert.strictEqual(service.isValidTestUrl('file:///C:/secrets.txt'), false)
assert.strictEqual(service.isValidTestUrl('not a url'), false)

console.log('WPS_MODEL_SERVICE=PASS')
