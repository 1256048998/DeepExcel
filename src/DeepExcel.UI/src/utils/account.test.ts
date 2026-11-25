import { describe, it, expect } from 'vitest'
import { normalizeAccountState, readAccountReply, readInviteRequired } from './account'

/**
 * 用户报的：登录 / 注册点了没反应。
 * 宿主其实回了 { type: 'error', payload: { message: '邮箱或密码不正确。' } }，
 * 面板把它当成了登录状态，于是既没登录成功、也没有任何提示。
 */

describe('readAccountReply', () => {
  it('surfaces the host error message instead of treating it as a sign-in state', () => {
    const reply = readAccountReply({ type: 'error', payload: { message: '邮箱或密码不正确。' } }, '登录失败')
    expect(reply).toEqual({ ok: false, error: '邮箱或密码不正确。' })
  })

  it('surfaces the invite-code rejection on register', () => {
    const reply = readAccountReply({ type: 'error', payload: { message: '需要有效的邀请码。' } }, '注册失败')
    expect(reply).toEqual({ ok: false, error: '需要有效的邀请码。' })
  })

  it('falls back to a generic message when the error has no text', () => {
    expect(readAccountReply({ type: 'error', payload: {} }, '登录失败')).toEqual({ ok: false, error: '登录失败' })
  })

  it('reports no reply as a network problem, not as silence', () => {
    const reply = readAccountReply(null, '登录失败')
    expect(reply.ok).toBe(false)
    if (!reply.ok) expect(reply.error).toMatch(/没有响应/)
  })

  it('accepts a real account status', () => {
    const reply = readAccountReply({
      type: 'account_status',
      payload: { state: 'signedin', email: 'a@b.com', server_url: 'https://x', mode: 'byok', entitlement: null },
    }, '登录失败')
    expect(reply.ok).toBe(true)
    if (reply.ok) expect(reply.status.email).toBe('a@b.com')
  })

  it('rejects a status payload without a known state', () => {
    expect(readAccountReply({ type: 'account_status', payload: { message: 'x' } }, '登录失败').ok).toBe(false)
  })

  it('keeps the restoring flag so the caller can wait instead of showing the sign-in page', () => {
    const reply = readAccountReply({ type: 'account_status', payload: { state: 'signedout', restoring: true } }, '')
    expect(reply.ok && reply.status.restoring).toBe(true)
  })
})

describe('normalizeAccountState', () => {
  it('maps the old signed_out spelling', () => {
    expect(normalizeAccountState('signed_out')).toBe('signedout')
  })
  it('passes known states through and rejects the rest', () => {
    expect(normalizeAccountState('offline')).toBe('offline')
    expect(normalizeAccountState('nope')).toBeNull()
    expect(normalizeAccountState(undefined)).toBeNull()
  })
})

describe('readInviteRequired', () => {
  it('reads the server answer', () => {
    expect(readInviteRequired({ type: 'account_server_meta', payload: { invite_required: true } })).toBe(true)
    expect(readInviteRequired({ type: 'account_server_meta', payload: { invite_required: false } })).toBe(false)
  })

  it('does not hide the invite field when the probe failed', () => {
    // 之前探测失败被读成 false：输入框消失，注册必然被服务器以缺邀请码拒绝
    expect(readInviteRequired({ type: 'error', payload: { message: '暂时连不上账号服务' } })).toBeNull()
    expect(readInviteRequired(null)).toBeNull()
  })
})
