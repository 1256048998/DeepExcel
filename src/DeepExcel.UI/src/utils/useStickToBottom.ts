import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { RefObject } from 'react'

/** 离底部不超过这么多像素就算「在底部」，继续自动跟随 */
export const STICK_THRESHOLD_PX = 60

export function isNearBottom(
  metrics: { scrollHeight: number; scrollTop: number; clientHeight: number },
  threshold = STICK_THRESHOLD_PX,
): boolean {
  return metrics.scrollHeight - metrics.scrollTop - metrics.clientHeight <= threshold
}

/**
 * 消息区的「贴底」滚动（Claude Code / 各家聊天面板的做法）：
 *   - 用户停在底部附近时，新内容到达就跟到底；用 instant 滚动并按帧合并，流式输出
 *     每秒几十次更新也只滚一次 / 帧，不会像 smooth 那样一路追不上、抖动
 *   - 用户往上翻超过 60px，就不再自动跟随，给出「回到底部」
 *   - contentKey 变化（新消息、流式增量）或内容区尺寸变化（展开 diff、图片加载）都会触发
 *   - forcePin 为 true 的那次更新（用户刚发了消息）无论在哪都回到底部
 */
export function useStickToBottom(
  containerRef: RefObject<HTMLElement>,
  contentRef: RefObject<HTMLElement>,
  contentKey: unknown,
  forcePin: boolean,
) {
  const pinnedRef = useRef(true)
  const [pinned, setPinned] = useState(true)
  const frameRef = useRef<number | null>(null)

  const scrollToBottom = useCallback(() => {
    if (frameRef.current !== null) return
    frameRef.current = requestAnimationFrame(() => {
      frameRef.current = null
      const target = containerRef.current
      if (target) target.scrollTop = target.scrollHeight
    })
  }, [containerRef])

  const updatePinned = useCallback((value: boolean) => {
    if (pinnedRef.current === value) return
    pinnedRef.current = value
    setPinned(value)
  }, [])

  // 用户滚动：离底部近就恢复跟随，远了就停止
  useEffect(() => {
    const el = containerRef.current
    if (!el) return
    const onScroll = () => updatePinned(isNearBottom(el))
    el.addEventListener('scroll', onScroll, { passive: true })
    return () => el.removeEventListener('scroll', onScroll)
  }, [containerRef, updatePinned])

  // 新内容到达
  useLayoutEffect(() => {
    if (forcePin) updatePinned(true)
    if (pinnedRef.current) scrollToBottom()
  }, [contentKey, forcePin, scrollToBottom, updatePinned])

  // 内容尺寸变化（Markdown 渲染完、展开详情、图片加载）
  useEffect(() => {
    const content = contentRef.current
    if (!content || typeof ResizeObserver === 'undefined') return
    const observer = new ResizeObserver(() => {
      if (pinnedRef.current) scrollToBottom()
    })
    observer.observe(content)
    return () => observer.disconnect()
  }, [contentRef, scrollToBottom])

  // 卸载（含 StrictMode 的模拟卸载）时取消未执行的帧，并清掉标记——否则「已排了一帧」
  // 的判断永远成立，之后再也不会滚动
  useEffect(() => () => {
    if (frameRef.current !== null) cancelAnimationFrame(frameRef.current)
    frameRef.current = null
  }, [])

  // 「回到底部」直接跳：smooth 动画途中的 scroll 事件会把「贴底」又判成否，按钮来回闪
  const jumpToBottom = useCallback(() => {
    updatePinned(true)
    const el = containerRef.current
    if (el) el.scrollTop = el.scrollHeight
  }, [containerRef, updatePinned])

  return { pinned, jumpToBottom }
}
