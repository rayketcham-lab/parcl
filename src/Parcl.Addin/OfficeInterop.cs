using System.Runtime.InteropServices;

namespace Microsoft.Office.Core
{
    /// <summary>
    /// Office Ribbon extensibility interface.
    /// Replaces dependency on office.dll PIA for self-contained builds.
    /// GUID: 000C0396-0000-0000-C000-000000000046
    /// </summary>
    [ComImport]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        string GetCustomUI(string RibbonID);
    }

    /// <summary>
    /// Represents the Ribbon UI instance passed to Ribbon_Load.
    /// GUID: 000C03A7-0000-0000-C000-000000000046
    /// </summary>
    [ComImport]
    [Guid("000C03A7-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonUI
    {
        [DispId(1)]
        void Invalidate();

        [DispId(2)]
        void InvalidateControl(string ControlID);

        [DispId(3)]
        void InvalidateControlMso(string ControlID);
    }

    /// <summary>
    /// Passed to ribbon callback methods to identify the control that triggered the action.
    /// GUID: 000C0395-0000-0000-C000-000000000046
    /// </summary>
    [ComImport]
    [Guid("000C0395-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonControl
    {
        [DispId(1)]
        string Id { [return: MarshalAs(UnmanagedType.BStr)] get; }

        [DispId(2)]
        object Context { [return: MarshalAs(UnmanagedType.IDispatch)] get; }

        [DispId(3)]
        string Tag { [return: MarshalAs(UnmanagedType.BStr)] get; }
    }
}
