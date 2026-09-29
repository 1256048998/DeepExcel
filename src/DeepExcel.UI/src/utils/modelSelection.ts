import type { ModelConfig } from '../types'
import type { ModelOption } from '../components/InputArea'
import type { AccountStatus } from '../components/AccountPanel'

/** 托管目录那一组在模型选择里的 provider 名，和宿主 MessageBridge.HostedProviderKey 一致 */
export const HOSTED_PROVIDER = 'hosted'

/**
 * 托管且服务端给了目录：模型选择就列目录里的模型（带每任务点数），不再看本地厂商。
 * 没有目录（或不是托管）返回 null，调用方退回本地厂商的模型——那是有目录之前的行为。
 * 路由模式和目录都是服务端给的，这里只是照着排出来。
 */
export function buildHostedModelOptions(account: AccountStatus | null): ModelOption[] | null {
  if (account?.mode !== 'hosted' || !account.hosted_models?.length) return null
  return account.hosted_models.map(m => ({
    provider: HOSTED_PROVIDER,
    providerDisplayName: 'DeepExcel 托管',
    model: m.model,
    label: m.label,
    isPrimary: m.default,
    points: m.points,
  }))
}

/**
 * 输入框模型下拉的选项，以及"现在到底能不能发消息"。
 *
 * 从 App.tsx 里提出来是为了能测——这两段逻辑出错都直接伤用户：下拉错了会让人
 * 用到不是自己选的模型，`setupNeeded` 错了会把一个其实能正常用的人拦在门外。
 * 和 C# 侧 TypeInference / PreviewBuilder 是同一个思路：把判断和 IO 分开，
 * 这样刁钻情况可以直接测。
 */

/**
 * 列出所有已配 API Key 的 provider 的模型，默认厂商置顶。
 *
 * 只要求 hasApiKey，不要求测试连接通过：否则默认厂商若没点过"测试连接"就会
 * 从下拉里消失，与"默认厂商的主模型是默认值"冲突。连接状态另外用圆点表示。
 */
export function buildModelOptions(config: ModelConfig | null): ModelOption[] {
  if (!config) return []

  const defaultProvider = config.defaultProvider || config.currentProvider
  const entries = Object.entries(config.providers ?? {})

  // 默认厂商置顶，其余保持 providers 的原有顺序。
  //
  // comparator 对两个非默认厂商返回 0，靠 Array.prototype.sort 的稳定性
  // （ES2019 起有保证）维持原序。这是刻意的：providers 的顺序本身就是用户在
  // 模型配置面板里排出来的，按字典序重排会把他排好的顺序打乱。
  entries.sort((a, b) => {
    if (a[0] === defaultProvider) return -1
    if (b[0] === defaultProvider) return 1
    return 0
  })

  const options: ModelOption[] = []
  for (const [key, provider] of entries) {
    if (!provider?.hasApiKey) continue
    for (const model of provider.models ?? []) {
      options.push({
        provider: key,
        providerDisplayName: provider.displayName,
        model,
        isPrimary: model === provider.defaultModel,
      })
    }
  }
  return options
}

/**
 * 该不该提示用户"先去配个模型"。
 *
 * 两个来源都必须先到齐，而且都 fail-open：
 *
 * - `config` 还是 null：配置没加载完。否则每次打开面板都会先闪一下引导，而
 *   多数用户早就配好了。
 * - `account` 还是 null：还不知道是不是托管模式。托管用户本地本来就没有
 *   API Key，这时候拦他们是纯误伤——而且 accountStatus 曾经只在用户打开过
 *   账号面板时才有值，也就是可能永远不到。
 *
 * 任一请求失败也停在 null，结果是不引导，退回改动前的行为。**宁可漏引导，
 * 也不要误拦一个其实能正常用的用户。**
 *
 * `account.mode` 由 `GET /api/v1/session/endpoint` 下发，这里只是应用它。
 * 不要改成从 plan / entitlement 反推路由——那会破坏出口路由契约。
 */
export function computeSetupNeeded(
  config: ModelConfig | null,
  account: AccountStatus | null,
  options: ModelOption[],
): boolean {
  if (config === null || account === null) return false
  if (account.mode === 'hosted') return false
  return options.length === 0
}
