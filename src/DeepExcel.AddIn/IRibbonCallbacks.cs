using System;
using System.Runtime.InteropServices;

namespace DeepExcel.AddIn
{
    /// <summary>
    /// Ribbon 回调 IDispatch 接口
    /// Excel 通过 IDispatch::Invoke 调用 Ribbon XML 中定义的 onLoad/onAction 回调方法名，
    /// 因此这些方法必须出现在 ThisAddIn 的默认 IDispatch 接口表中。
    /// 参数类型用 object（对应 COM 的 IDispatch*），由 Office 传入 IRibbonControl/IRibbonUI 对象。
    /// </summary>
    [ComVisible(true)]
    [Guid("C3D4E5F6-A7B8-4C5D-9A7F-1C2D3E4F5A6B")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonCallbacks
    {
        [DispId(1)]
        void OnRibbonLoad(object ribbon);

        [DispId(2)]
        void OnTogglePanel(object control);

        [DispId(3)]
        void OnShowHelp(object control);

        [DispId(4)]
        void OnShowModelConfig(object control);
    }
}
