// 附件 chip 的展示：按扩展名分类给图标，图片用上传时读到的数据做缩略图。
// 宿主的 list_attachments 只回文件名和大小，缩略图只对本次面板里上传的图片有；
// 其它情况退回类型图标。纯函数，有测试：attachments.test.ts

export type AttachmentKind = 'image' | 'sheet' | 'pdf' | 'doc' | 'text'

const KIND_BY_EXT: Record<string, AttachmentKind> = {
  png: 'image', jpg: 'image', jpeg: 'image', gif: 'image', bmp: 'image', webp: 'image', svg: 'image', tiff: 'image',
  xlsx: 'sheet', xls: 'sheet', csv: 'sheet',
  pdf: 'pdf',
  doc: 'doc', docx: 'doc', ppt: 'doc', pptx: 'doc',
}

export function attachmentKind(fileName: string): AttachmentKind {
  const dot = fileName.lastIndexOf('.')
  const ext = dot >= 0 ? fileName.slice(dot + 1).toLowerCase() : ''
  return KIND_BY_EXT[ext] ?? 'text'
}

// 缩略图只留小于这个大小的图：数据 URL 常驻内存，大图没必要
export const THUMBNAIL_MAX_BYTES = 8 * 1024 * 1024

export function keepThumbnail(fileName: string, size: number): boolean {
  return attachmentKind(fileName) === 'image' && size <= THUMBNAIL_MAX_BYTES
}

export function formatSize(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) return ''
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`
}

// 剪贴板里的截图文件名一律叫 image.png，同名上传会互相覆盖；按时间起名
export function pastedImageName(type: string, now: Date = new Date()): string {
  const ext = ({ 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/webp': 'webp', 'image/bmp': 'bmp' } as Record<string, string>)[type] ?? 'png'
  const p = (n: number) => String(n).padStart(2, '0')
  const stamp = `${now.getFullYear()}${p(now.getMonth() + 1)}${p(now.getDate())}-${p(now.getHours())}${p(now.getMinutes())}${p(now.getSeconds())}`
  return `截图-${stamp}.${ext}`
}
