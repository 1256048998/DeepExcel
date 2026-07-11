using System;
using System.Runtime.InteropServices;

namespace Microsoft.Office.Core
{
    [ComVisible(true)]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        string GetCustomUI(string RibbonID);
    }

    [ComVisible(true)]
    [Guid("000C03A7-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonUI
    {
        void Invalidate();
        void InvalidateControl(string ControlID);
    }

    [ComVisible(true)]
    [Guid("000C0395-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonControl
    {
        string Id { get; }
        object Context { get; }
        string Tag { get; }
    }
}
