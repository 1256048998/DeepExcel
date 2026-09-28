import { describe, expect, it } from 'vitest'
import { attachmentKind, formatSize, keepThumbnail, pastedImageName, THUMBNAIL_MAX_BYTES } from './attachments'

describe('attachment chips', () => {
  it('classifies by extension, case-insensitively', () => {
    expect(attachmentKind('发票.PNG')).toBe('image')
    expect(attachmentKind('销售明细.xlsx')).toBe('sheet')
    expect(attachmentKind('a.csv')).toBe('sheet')
    expect(attachmentKind('合同.pdf')).toBe('pdf')
    expect(attachmentKind('说明.docx')).toBe('doc')
    expect(attachmentKind('notes.md')).toBe('text')
    expect(attachmentKind('noext')).toBe('text')
  })

  it('keeps thumbnails only for reasonably sized images', () => {
    expect(keepThumbnail('a.jpg', 1000)).toBe(true)
    expect(keepThumbnail('a.jpg', THUMBNAIL_MAX_BYTES + 1)).toBe(false)
    expect(keepThumbnail('a.xlsx', 1000)).toBe(false)
  })

  it('formats sizes', () => {
    expect(formatSize(512)).toBe('512 B')
    expect(formatSize(20 * 1024)).toBe('20 KB')
    expect(formatSize(3.5 * 1024 * 1024)).toBe('3.5 MB')
  })

  it('names pasted screenshots by time so they do not overwrite each other', () => {
    const t = new Date(2026, 8, 28, 9, 5, 7)
    expect(pastedImageName('image/png', t)).toBe('截图-20260928-090507.png')
    expect(pastedImageName('image/jpeg', t)).toBe('截图-20260928-090507.jpg')
  })
})
