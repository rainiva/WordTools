using System.Runtime.InteropServices;

namespace WordTools.Interop
{
    [ComImport]
    [Guid("00072DB7-00A0-4211-A370-6E7B0AE64EA1")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IRequestComAddInAutomationService
    {
        [return: MarshalAs(UnmanagedType.IDispatch)]
        object GetComAddInAutomationService();
    }
}
