// config.json 由 Excel 插件和 WPS 共用。WPS 的 ConfigStore 按固定字段重建配置，
// 以前在 WPS 里保存一次，就会把 Excel 端独有的字段（Update 更新设置、托管模型 HostedModel）抹掉。
// 用法：node scripts/test-wps-config-store.js
const assert = require('assert')
const fs = require('fs')
const os = require('os')
const path = require('path')

const ConfigStore = require('../src/DeepExcel.Wps/config-store.js')

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'deepexcel-config-'))
const filePath = path.join(dir, 'config.json')
const fakeCredentials = { get: () => '', set: () => true, remove: () => true }

try {
  fs.writeFileSync(filePath, JSON.stringify({
    CurrentProvider: 'deepseek',
    CurrentModel: 'deepseek-v4-pro',
    HostedModel: 'deepseek-v4-flash',
    Update: { FeedUrl: 'https://updates.example/api/v1/updates/latest', Channel: 'beta' },
    General: { MaxTurns: 30 },
  }), 'utf8')

  const store = new ConfigStore({ dir, filePath, credentials: fakeCredentials })
  const cfg = store.load()
  assert.strictEqual(cfg.HostedModel, 'deepseek-v4-flash')
  assert.deepStrictEqual(cfg.Update, { FeedUrl: 'https://updates.example/api/v1/updates/latest', Channel: 'beta' })
  assert.strictEqual(cfg.General.MaxTurns, 30)
  assert.strictEqual(cfg.CurrentProvider, 'deepseek')

  assert.strictEqual(store.save(), true)
  const written = JSON.parse(fs.readFileSync(filePath, 'utf8'))
  assert.strictEqual(written.HostedModel, 'deepseek-v4-flash', 'HostedModel must survive a WPS save')
  assert.strictEqual(written.Update.Channel, 'beta', 'Update settings must survive a WPS save')
  // 认识的字段仍按规范重建（例如 providers 补全、catalog 版本升级）
  assert.ok(written.Providers.deepseek)
  assert.strictEqual(written.ModelCatalogVersion, ConfigStore.CURRENT_MODEL_CATALOG_VERSION)

  // 大小写不同的已知字段不能被当成未知字段带两份
  fs.writeFileSync(filePath, JSON.stringify({ currentModel: 'deepseek-v4-pro', currentProvider: 'deepseek' }), 'utf8')
  const again = new ConfigStore({ dir, filePath, credentials: fakeCredentials }).load()
  assert.ok(!('currentModel' in again), 'known keys are normalised, not passed through')

  console.log('test-wps-config-store: all passed')
} finally {
  fs.rmSync(dir, { recursive: true, force: true })
}
