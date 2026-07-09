// src/DeepExcel.Wps/jsa-executor.js
// JSA 宏执行（对应 C# 端 VBAExecutor.cs）
// 职责：执行 WPS JS 宏代码（JSA），替代 VBA
//
// ★ WPS JSA 与 VBA 的差异：
// - JSA 语法为 ES6（let/const/箭头函数/模板字符串）
// - 对象模型与 VBA 一致（Application/Workbook/Worksheet/Range）
// - 入口：通过 wps.Application.Run(macroName) 调用已注册的 JSA 函数
//
// ★ 执行策略：
// 1. 把 AI 写的 JSA 代码包装为函数
// 2. 通过 Application.Run 执行
// 3. 捕获异常返回给 AI

class JsaExecutor {
  /**
   * 执行 JSA 代码
   * @param {string} code JSA 代码（ES6 语法）
   * @returns {Promise<{success: boolean, data: any, error: string}>}
   */
  async execute(code) {
    if (!code || !code.trim()) {
      return { success: false, data: {}, error: 'JSA 代码为空' }
    }

    console.log(`[JsaExecutor] Execute, code length=${code.length}`)
    console.log(`[JsaExecutor] Code preview: ${code.slice(0, 200)}...`)

    try {
      // ★ 方式 1：通过 Application.Run 执行预定义宏
      // ★ 方式 2：通过 wps.Application.Run 传字符串执行
      // 这里用 Application.Run 包装为匿名函数调用
      const macroName = '__deepexcel_temp_' + Date.now()
      const wrappedCode = `
        function ${macroName}() {
          try {
            ${code}
            return { success: true }
          } catch (e) {
            return { success: false, error: e.message }
          }
        }
      `

      // ★ 先注入宏定义到 VBA/JSA 工程中（如果 WPS 支持）
      // 实际执行方式取决于 WPS 版本和 API 能力
      const app = wps.Application
      let result
      try {
        // ★ 尝试方式 1：直接通过 Application.Macro.JSEval 执行（如果 WPS 支持）
        if (app.Macro && app.Macro.JSEval) {
          result = await app.Macro.JSEval(code)
        } else {
          // ★ 方式 2：通过 Application.Run 执行
          // 需要先把宏定义注入到工作簿的 JSA 模块中
          result = await this._injectAndRun(app, macroName, wrappedCode)
        }
      } catch (execErr) {
        return {
          success: false,
          data: {},
          error: `JSA 执行失败: ${execErr.message}`,
        }
      }

      // ★ 统一结果格式
      if (result && typeof result === 'object' && 'success' in result) {
        return { success: result.success, data: result.data || {}, error: result.error || '' }
      }
      return { success: true, data: { result }, error: '' }
    } catch (err) {
      console.error('[JsaExecutor] Execute FAILED:', err)
      return {
        success: false,
        data: {},
        error: `JSA 执行异常: ${err.message}`,
      }
    }
  }

  /**
   * 注入宏定义并执行（通过 VBProject 或 JSA 模块）
   * ★ 这是 WPS 下的最佳尝试，具体能力取决于 WPS 版本和是否启用宏
   */
  async _injectAndRun(app, macroName, wrappedCode) {
    try {
      // ★ 尝试通过 VBProject 注入（如果 WPS 安装了 VBA 插件）
      const wb = app.ActiveWorkbook
      if (wb && wb.VBProject) {
        const vbComponents = wb.VBProject.VBComponents
        const moduleName = 'DeepExcelTemp'
        let module
        try {
          module = vbComponents(moduleName)
        } catch (e) {
          // 模块不存在，创建新的
          module = vbComponents.Add(1) // vbext_ct_StdModule=1
          module.Name = moduleName
        }
        // 清空旧代码并注入新代码
        const codeModule = module.CodeModule
        const lineCount = codeModule.CountOfLines
        if (lineCount > 0) {
          codeModule.DeleteLines(1, lineCount)
        }
        codeModule.AddFromString(wrappedCode)
        // 执行宏
        return app.Run(macroName)
      }
    } catch (e) {
      console.warn('[JsaExecutor] VBProject injection failed:', e.message)
    }

    // ★ VBProject 不可用时，尝试通过 JSA 直接 eval
    try {
      // eslint-disable-next-line no-new-func
      const fn = new Function(`return (function() { ${wrappedCode} })()`)
      return fn()
    } catch (evalErr) {
      throw new Error(`JSA 执行失败（VBProject 和 eval 都不可用）: ${evalErr.message}`)
    }
  }
}

module.exports = JsaExecutor
