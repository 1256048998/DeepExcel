// src/DeepExcel.Wps/credential-store.js
// API Key 存取（对应 C# 端 Security/SecurityManager.cs）
//
// 存储位置与格式与 Excel 加载项完全一致：
//   %APPDATA%\DeepExcel\credentials\key_<provider>.crypt，内容 = DPAPI 密文的 base64
// 因此同一台机器上 Excel 和 WPS 共用一份凭据，用户只需配置一次 API Key。
//
// Node 没有 DPAPI 绑定，加解密委托给 dpapi-cli.py（加载项本来就依赖 Python sidecar）。
// has / delete 只看文件系统，不需要起 Python 进程。

const { execFileSync } = require('child_process')
const fs = require('fs')
const path = require('path')

// ★ provider key 会拼进文件名，必须白名单校验，防止 ../ 之类的路径穿越
const PROVIDER_PATTERN = /^[A-Za-z0-9_-]{1,64}$/

class CredentialStore {
  constructor(options = {}) {
    this.dir = options.dir ||
      path.join(process.env.APPDATA || '', 'DeepExcel', 'credentials')
    this.pythonPath = options.pythonPath || 'python'
    this.scriptPath = options.scriptPath || path.join(__dirname, 'dpapi-cli.py')
    this.timeoutMs = options.timeoutMs || 8000
  }

  /** 更新 Python 解释器路径（main.js 拿到 SidecarHost 探测结果后回填） */
  setPythonPath(pythonPath) {
    if (pythonPath) this.pythonPath = pythonPath
  }

  pathFor(provider) {
    if (!PROVIDER_PATTERN.test(provider || '')) return null
    return path.join(this.dir, `key_${provider}.crypt`)
  }

  has(provider) {
    const file = this.pathFor(provider)
    if (!file) return false
    try {
      return fs.existsSync(file) && fs.statSync(file).size > 0
    } catch (e) {
      return false
    }
  }

  /** 返回明文 key；未配置或解密失败返回空字符串（fail-closed，与 C# 一致） */
  get(provider) {
    if (!this.has(provider)) return ''
    try {
      const out = execFileSync(this.pythonPath, [this.scriptPath, 'get', provider], {
        encoding: 'utf8',
        timeout: this.timeoutMs,
        windowsHide: true,
      })
      return (out || '').trim()
    } catch (e) {
      console.error('[CredentialStore] get failed:', e.message)
      return ''
    }
  }

  set(provider, apiKey) {
    if (!PROVIDER_PATTERN.test(provider || '') || !apiKey) return false
    try {
      execFileSync(this.pythonPath, [this.scriptPath, 'set', provider], {
        input: apiKey,
        encoding: 'utf8',
        timeout: this.timeoutMs,
        windowsHide: true,
      })
      return this.has(provider)
    } catch (e) {
      console.error('[CredentialStore] set failed:', e.message)
      return false
    }
  }

  remove(provider) {
    const file = this.pathFor(provider)
    if (!file) return false
    try {
      if (fs.existsSync(file)) fs.unlinkSync(file)
      return true
    } catch (e) {
      console.error('[CredentialStore] remove failed:', e.message)
      return false
    }
  }

  /** 脱敏预览，规则与 C# SecurityManager.MaskApiKey 一致 */
  static mask(key) {
    if (!key) return ''
    if (key.length < 8) return '***'
    return `${key.slice(0, 4)}***...${key.slice(-3)}`
  }
}

module.exports = CredentialStore
