/**
 * 模型配置弹窗里「正在看哪家厂商」的决定。
 *
 * 这件事之前写在一个 `[config]` effect 里：config 每变一次就把选中厂商重置成
 * config.currentProvider。可 config 变化的来源远不止"打开弹窗"——导入 / 添加 /
 * 删除 / 拖拽模型（persistModels 回写 config）、测试连接、删 Key、设默认厂商
 * （都会 loadConfig）都会改它。于是用户在别的厂商下点「获取模型」→「添加」，
 * 界面就跳回当前厂商（没保存过时就是 Anthropic）。
 *
 * 规则：只有弹窗刚打开（resync）或者当前选中的厂商已经不存在时，才跟随
 * currentProvider；其他时候用户选了谁就是谁。
 */
export interface ProviderSyncConfig {
  currentProvider: string
  providers: Record<string, unknown>
}

export function pickSelectedProvider(
  current: string,
  config: ProviderSyncConfig,
  resync: boolean,
): string {
  if (!resync && current && Object.prototype.hasOwnProperty.call(config.providers, current)) {
    return current
  }
  return config.currentProvider
}

/**
 * 表单（API Key 输入框、Base URL、最大轮数）要不要整体重置成已保存的值。
 *
 * 只在换了厂商（或弹窗刚打开，调用方把 formProvider 清空）时重置。同一家厂商
 * 的 config 刷新——比如刚导入完模型——不能把用户正在输入、还没保存的 Key 和
 * Base URL 冲掉：「获取模型」本身不落盘 Key，冲掉之后「保存并应用」就保存了个寂寞。
 */
export function shouldResetForm(formProvider: string, selected: string): boolean {
  return formProvider !== selected
}
