// 托管模式下每个任务扣几点：服务端在 entitlement.model_weights 里按模型名前缀给出，
// 面板只负责对上号显示，扣多少始终以服务端为准（server/app/proxy/metering.py）。
// 对不上任何前缀时用服务端给的 model_weight_default（没定价的模型按最贵的扣），
// 服务端没给就不显示，不猜。纯函数，有测试：points.test.ts

export type ModelWeights = Record<string, number>

export function pointsForModel(
  model: string, weights: ModelWeights | null | undefined, fallback?: number | null,
): number | null {
  if (!weights || !model) return null
  const name = model.toLowerCase()
  // 最长前缀优先，和服务端一致：claude-sonnet-5 的单独定价盖过 claude-sonnet
  const prefix = Object.keys(weights)
    .filter(p => name.startsWith(p.toLowerCase()))
    .sort((a, b) => b.length - a.length)[0]
  if (prefix !== undefined) return weights[prefix]
  return fallback ?? null
}

export function pointsLabel(points: number): string {
  return `${points} 点/任务`
}
