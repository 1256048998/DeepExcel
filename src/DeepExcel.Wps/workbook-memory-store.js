// src/DeepExcel.Wps/workbook-memory-store.js
// 工作簿记忆的面板入口（对应 C# 端 Collaboration/WorkbookMemoryStore.cs）
//
// 记忆由侧车维护（workbook_memory.py）；这里只读写同一组文件，目录口径三端一致：
//   %LOCALAPPDATA%\DeepExcel\workbooks\<sha256(workbookKey) 前 16 字节 hex>\NOTES.md / history.jsonl / meta.json
// 所以同一个文件在 Excel 里记下的偏好和禁区，在 WPS 里一样生效、一样能改。

const crypto = require('crypto')
const fs = require('fs')
const path = require('path')

const MAX_NOTES_CHARS = 6000

function rootDir() {
  if (process.env.DEEPEXCEL_MEMORY_DIR) return process.env.DEEPEXCEL_MEMORY_DIR
  return path.join(process.env.LOCALAPPDATA || '', 'DeepExcel', 'workbooks')
}

function keyHash(workbookKey) {
  return crypto.createHash('sha256').update(String(workbookKey), 'utf8').digest().slice(0, 16).toString('hex')
}

/** 没保存过的工作簿（key 不是路径）没有记忆 */
function isPersistentKey(workbookKey) {
  return typeof workbookKey === 'string' && (workbookKey.indexOf('\\') >= 0 || workbookKey.indexOf('/') >= 0)
}

function dirFor(workbookKey) {
  return path.join(rootDir(), keyHash(workbookKey))
}

function readText(file) {
  try {
    return fs.existsSync(file) ? fs.readFileSync(file, 'utf8') : ''
  } catch (e) {
    return ''
  }
}

function countLines(file) {
  const text = readText(file)
  return text ? text.split('\n').filter(line => line.trim()).length : 0
}

function describe(workbookKey, workbookName, error) {
  if (!isPersistentKey(workbookKey)) {
    return { available: false, workbookName: workbookName || '', notes: '', historyCount: 0, error: error || null }
  }
  const dir = dirFor(workbookKey)
  return {
    available: true,
    workbookName: workbookName || '',
    notes: readText(path.join(dir, 'NOTES.md')),
    historyCount: countLines(path.join(dir, 'history.jsonl')),
    error: error || null,
  }
}

/** 保存用户改过的记忆；不允许时返回原因 */
function save(workbookKey, workbookName, notes) {
  if (!isPersistentKey(workbookKey)) return '工作簿还没保存过，保存后才有记忆'
  const text = String(notes || '').replace(/\r\n/g, '\n').trim()
  if (text.length > MAX_NOTES_CHARS) return '记忆太长（' + text.length + ' 字），上限 ' + MAX_NOTES_CHARS + ' 字'
  const dir = dirFor(workbookKey)
  fs.mkdirSync(dir, { recursive: true })
  const file = path.join(dir, 'NOTES.md')
  if (!text) {
    try { fs.unlinkSync(file) } catch (e) { /* 本来就没有 */ }
  } else {
    const temp = file + '.tmp'
    fs.writeFileSync(temp, text + '\n', 'utf8')
    fs.renameSync(temp, file)
  }
  try {
    fs.writeFileSync(path.join(dir, 'meta.json'), JSON.stringify({
      workbook_key: workbookKey, workbook_name: workbookName || '', updated_at: Math.floor(Date.now() / 1000),
    }), 'utf8')
  } catch (e) { /* meta 只是面板列表用 */ }
  return null
}

function clear(workbookKey) {
  if (!isPersistentKey(workbookKey)) return
  const dir = dirFor(workbookKey)
  for (const name of ['NOTES.md', 'history.jsonl', 'meta.json']) {
    try { fs.unlinkSync(path.join(dir, name)) } catch (e) { /* 本来就没有 */ }
  }
  try { fs.rmdirSync(dir) } catch (e) { /* 目录里还有别的东西，或本来就没有 */ }
}

module.exports = { MAX_NOTES_CHARS, keyHash, isPersistentKey, dirFor, describe, save, clear }
