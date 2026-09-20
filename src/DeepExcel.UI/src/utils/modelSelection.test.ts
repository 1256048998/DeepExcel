import { describe, it, expect } from 'vitest'
import { buildModelOptions, computeSetupNeeded } from './modelSelection'
import type { ModelConfig } from '../types'
import type { AccountStatus } from '../components/AccountPanel'

/**
 * 面板上唯一两段"算错就直接伤用户"的纯逻辑。
 *
 * computeSetupNeeded 的 fail-open 尤其值得锁：它第一版就把托管用户拦在了门外，
 * 因为那些人本地本来就没有 API Key，而路由模式是异步到达的。写的时候很自然，
 * 错了也没有任何东西会报警——受影响的恰恰是付费那批人。
 */

function provider(over: Partial<{
  displayName: string; models: string[]; defaultModel: string; hasApiKey: boolean
}> = {}) {
  return {
    type: 'anthropic',
    displayName: over.displayName ?? 'Claude',
    baseUrl: 'https://api.anthropic.com',
    models: over.models ?? ['m1', 'm2'],
    defaultModel: over.defaultModel ?? 'm1',
    hasApiKey: over.hasApiKey ?? true,
    supportsVision: false,
    connected: true,
  } as any
}

function config(providers: Record<string, any>, over: Partial<ModelConfig> = {}): ModelConfig {
  return {
    currentProvider: over.currentProvider ?? 'anthropic',
    currentModel: over.currentModel ?? 'm1',
    defaultProvider: over.defaultProvider ?? 'anthropic',
    providers,
    general: {} as any,
    ...over,
  } as ModelConfig
}

const signedOut: AccountStatus = {
  state: 'signedout', server_url: null, email: null, mode: null, entitlement: null,
}
const hosted: AccountStatus = {
  state: 'signedin', server_url: 'https://x', email: 'a@b.c', mode: 'hosted', entitlement: null,
}
const byok: AccountStatus = {
  state: 'signedin', server_url: 'https://x', email: 'a@b.c', mode: 'byok', entitlement: null,
}

describe('buildModelOptions', () => {
  it('returns nothing before the config has loaded', () => {
    expect(buildModelOptions(null)).toEqual([])
  })

  it('lists every model of every provider that has a key', () => {
    const options = buildModelOptions(config({
      anthropic: provider({ models: ['a1', 'a2'] }),
      deepseek: provider({ displayName: 'DeepSeek', models: ['d1'] }),
    }))

    expect(options.map(o => `${o.provider}::${o.model}`))
      .toEqual(['anthropic::a1', 'anthropic::a2', 'deepseek::d1'])
  })

  it('skips providers with no API key', () => {
    // Not "no successful test" -- just no key. A provider the user configured
    // but never pressed "test connection" on must still be usable, otherwise
    // the default provider vanishes from the dropdown.
    const options = buildModelOptions(config({
      anthropic: provider({ models: ['a1'], hasApiKey: false }),
      deepseek: provider({ displayName: 'DeepSeek', models: ['d1'] }),
    }))

    expect(options.map(o => o.provider)).toEqual(['deepseek'])
  })

  it('puts the default provider first', () => {
    const options = buildModelOptions(config({
      zeta: provider({ displayName: 'Zeta', models: ['z1'] }),
      anthropic: provider({ models: ['a1'] }),
    }, { defaultProvider: 'anthropic' }))

    expect(options[0].provider).toBe('anthropic')
  })

  it('keeps the remaining providers in their configured order', () => {
    // The order in `providers` is the order the user dragged them into in the
    // model panel. Re-sorting alphabetically would silently undo that.
    const options = buildModelOptions(config({
      zeta: provider({ displayName: 'Zeta', models: ['z1'] }),
      alpha: provider({ displayName: 'Alpha', models: ['x1'] }),
      anthropic: provider({ models: ['a1'] }),
    }, { defaultProvider: 'anthropic' }))

    expect(options.map(o => o.provider)).toEqual(['anthropic', 'zeta', 'alpha'])
  })

  it('marks only the provider default as primary', () => {
    const options = buildModelOptions(config({
      anthropic: provider({ models: ['a1', 'a2'], defaultModel: 'a2' }),
    }))

    expect(options.filter(o => o.isPrimary).map(o => o.model)).toEqual(['a2'])
  })

  it('falls back to currentProvider when no default is set', () => {
    const options = buildModelOptions(config({
      zeta: provider({ displayName: 'Zeta', models: ['z1'] }),
      deepseek: provider({ displayName: 'DeepSeek', models: ['d1'] }),
    }, { defaultProvider: '', currentProvider: 'deepseek' }))

    expect(options[0].provider).toBe('deepseek')
  })

  it('survives a provider with no models', () => {
    const options = buildModelOptions(config({
      anthropic: provider({ models: [] }),
    }))

    expect(options).toEqual([])
  })
})

describe('computeSetupNeeded', () => {
  it('prompts when signed out with no key configured', () => {
    expect(computeSetupNeeded(config({}), signedOut, [])).toBe(true)
  })

  it('stays quiet once a model is available', () => {
    const options = buildModelOptions(config({ anthropic: provider() }))
    expect(computeSetupNeeded(config({ anthropic: provider() }), signedOut, options)).toBe(false)
  })

  it('never prompts a hosted user', () => {
    // The whole point. Hosted users have no local key by design -- model
    // traffic goes through the server proxy. Prompting them to "configure a
    // provider" is wrong advice, and blocking their message is worse.
    expect(computeSetupNeeded(config({}), hosted, [])).toBe(false)
  })

  it('does prompt a byok user with no key', () => {
    expect(computeSetupNeeded(config({}), byok, [])).toBe(true)
  })

  it('stays quiet while the model config is still loading', () => {
    // Otherwise the notice flashes on every panel open before the config
    // arrives, for users who configured everything long ago.
    expect(computeSetupNeeded(null, signedOut, [])).toBe(false)
  })

  it('stays quiet while the account status is still loading', () => {
    // This is the regression that shipped and had to be fixed: account status
    // arrives asynchronously -- and used to only arrive if the user happened
    // to open the account panel -- so treating "not yet known" as "not hosted"
    // locked hosted users out of the product.
    expect(computeSetupNeeded(config({}), null, [])).toBe(false)
  })

  it('stays quiet when both are still unknown', () => {
    expect(computeSetupNeeded(null, null, [])).toBe(false)
  })
})
