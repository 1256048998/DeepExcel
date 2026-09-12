using System;
using System.Runtime.InteropServices;

namespace Extensibility
{
    // ★ [ComImport] 必须有：声明这是"从 COM 导入的接口"，CLR 会按 COM vtable 布局生成 CCW。
    // 若改用 [ComVisible(true)]，CLR 会将其视为"托管接口导出给 COM"，vtable 布局与 Excel 期望的
    // IDTExtensibility2 不一致，导致 QI 成功但 OnConnection 调用时 vtable 偏移错误，静默失败。
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom);
        void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom);
        void OnAddInsUpdate(ref Array custom);
        void OnStartupComplete(ref Array custom);
        void OnBeginShutdown(ref Array custom);
    }

    [ComVisible(false)]
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4,
        ext_cm_UISetup = 5
    }

    [ComVisible(false)]
    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UISetupCompleted = 2,
        ext_dm_SolutionClosed = 3
    }
}
