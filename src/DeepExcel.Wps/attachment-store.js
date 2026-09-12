// src/DeepExcel.Wps/attachment-store.js
// 附件管理（对应 C# 端 WorkbookSession 的 AddAttachment / RemoveAttachment / GetAttachmentList）
//
// 存储在 %LOCALAPPDATA%\DeepExcel\Attachments\<workbookKey 哈希>\，按工作簿隔离。
// 注意：C# 端目录名用的是 .NET String.GetHashCode()，JS 无法复现，所以这里用 SHA256 前 4 字节。
// 附件是"本次对话带进来的临时素材"，不像 config / history 那样需要跨宿主共享，
// 各自独立目录反而更干净（同一台机器两个宿主不会互相删文件）。

const crypto = require('crypto')
const fs = require('fs')
const path = require('path')

const MAX_ATTACHMENT_SIZE = 10 * 1024 * 1024          // 单文件 10MB
const MAX_BASE64_LENGTH = 14 * 1024 * 1024            // 10MB 原文约 13.4MB base64

/** 允许的扩展名白名单，与 C# MessageBridge.HandleUploadAttachment 保持一致 */
const ALLOWED_EXTENSIONS = new Set([
  '.xlsx', '.xls', '.csv', '.txt', '.md', '.json', '.xml',
  '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.webp', '.tiff',
  '.pdf', '.doc', '.docx', '.ppt', '.pptx',
  '.py', '.js', '.ts', '.sql', '.html', '.css',
])

class AttachmentStore {
  constructor(workbookKey, options = {}) {
    const hash = crypto.createHash('sha256').update(String(workbookKey), 'utf8')
      .digest().slice(0, 4).toString('hex').toUpperCase()
    this.dir = options.dir ||
      path.join(process.env.LOCALAPPDATA || '', 'DeepExcel', 'Attachments', hash)
    // fileName -> 绝对路径
    this.files = new Map()
    try {
      fs.mkdirSync(this.dir, { recursive: true })
    } catch (e) {
      console.warn('[AttachmentStore] mkdir failed:', e.message)
    }
  }

  /**
   * 保存上传的附件，返回 { fileName, filePath, size }。
   * 校验：文件名安全化 → 扩展名白名单 → base64 合法性 → 大小上限 → 重名加序号。
   * 校验不过抛 Error，由调用方转成前端错误消息。
   */
  add(fileName, fileBase64) {
    const safeBase = path.basename(String(fileName || '')) || 'unnamed'
    const ext = path.extname(safeBase).toLowerCase()
    if (!ALLOWED_EXTENSIONS.has(ext)) {
      throw new Error(`不支持的文件类型: ${ext}（仅允许文档、图片、表格、代码等常用格式）`)
    }
    if (typeof fileBase64 !== 'string' || fileBase64.length === 0) {
      throw new Error('缺少 file_base64 参数')
    }
    if (fileBase64.length > MAX_BASE64_LENGTH) {
      throw new Error('附件大小超过限制（10MB），请压缩或选择更小的文件')
    }

    let bytes
    try {
      bytes = Buffer.from(fileBase64, 'base64')
    } catch (e) {
      throw new Error('文件内容不是有效的 base64 编码')
    }
    if (!bytes || bytes.length === 0) throw new Error('文件内容不是有效的 base64 编码')
    if (bytes.length > MAX_ATTACHMENT_SIZE) {
      throw new Error(`文件过大 (${Math.round(bytes.length / 1024 / 1024)}MB)，最大支持 10MB`)
    }

    let finalName = safeBase
    let target = path.join(this.dir, finalName)
    if (fs.existsSync(target)) {
      const stem = path.basename(safeBase, ext)
      let n = 1
      while (fs.existsSync(path.join(this.dir, `${stem}_${n}${ext}`))) n++
      finalName = `${stem}_${n}${ext}`
      target = path.join(this.dir, finalName)
    }

    fs.writeFileSync(target, bytes)
    this.files.set(finalName, target)
    return { fileName: finalName, filePath: target, size: bytes.length }
  }

  remove(fileName) {
    if (!this.files.has(fileName)) return false
    const target = this.files.get(fileName)
    try {
      if (fs.existsSync(target)) fs.unlinkSync(target)
    } catch (e) {
      console.warn('[AttachmentStore] remove failed:', e.message)
    }
    this.files.delete(fileName)
    return true
  }

  /** 前端列表用：[{ fileName, size }] */
  list() {
    const result = []
    for (const [name, target] of this.files) {
      let size = 0
      try {
        if (fs.existsSync(target)) size = fs.statSync(target).size
      } catch (e) { /* 读不到大小就按 0 */ }
      result.push({ fileName: name, size })
    }
    return result
  }

  /** 发给 sidecar 的上下文用：[{ name, size, path }]，agent 靠 path 自己去读文件 */
  contextList() {
    return this.list().map(item => ({
      name: item.fileName,
      size: item.size,
      path: this.files.get(item.fileName),
    }))
  }
}

module.exports = AttachmentStore
module.exports.ALLOWED_EXTENSIONS = ALLOWED_EXTENSIONS
