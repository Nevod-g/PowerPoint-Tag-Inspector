using System.Runtime.InteropServices;

namespace PowerPointTagInspector.Services;

/// <summary>
/// Provides GetActiveObject functionality for .NET 8+ where Marshal.GetActiveObject is unavailable.
/// </summary>
internal static class ComInteropHelper
{
    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(
        ref Guid rclsid,
        nint pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

    [DllImport("ole32.dll")]
    private static extern int CLSIDFromProgID(
        [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
        out Guid pclsid);

    /// <summary>
    /// Gets the running COM object registered with the specified ProgID.
    /// Equivalent to Marshal.GetActiveObject which was removed in .NET Core/.NET 5+.
    /// </summary>
    /// <exception cref="COMException">Thrown when no running instance is found.</exception>
    public static object GetActiveObject(string progId)
    {
        ArgumentNullException.ThrowIfNull(progId);

        int hr = CLSIDFromProgID(progId, out Guid clsid);

        if (hr < 0)
        {
            Marshal.ThrowExceptionForHR(hr);
        }

        GetActiveObject(ref clsid, nint.Zero, out object obj);

        return obj;
    }
}
