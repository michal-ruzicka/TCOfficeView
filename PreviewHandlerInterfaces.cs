using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace TCOffice.Host;

// ---------------------------------------------------------------------------
// Preview Handler COM interfaces
// https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-ipreviewhandler
// ---------------------------------------------------------------------------

[StructLayout(LayoutKind.Sequential)]
public struct RECT
{
    public int left, top, right, bottom;
    public RECT(int l, int t, int r, int b) { left = l; top = t; right = r; bottom = b; }
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("8895b1c6-b41f-4c1c-a562-0d564250836f")]
public interface IPreviewHandler
{
    void SetWindow(IntPtr hwnd, ref RECT rect);
    void SetRect(ref RECT rect);
    void DoPreview();
    void Unload();
    void SetFocus();
    void QueryFocus(out IntPtr phwnd);
    [PreserveSig]
    int TranslateAccelerator(ref System.Windows.Forms.Message pmsg);
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("b7d14566-0509-4cce-a71f-0a554233bd9b")]
public interface IInitializeWithFile
{
    void Initialize([MarshalAs(UnmanagedType.LPWStr)] string pszFilePath, uint grfMode);
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("b824b49d-22ac-4132-ac65-c0a17b6a3b9c")]
public interface IInitializeWithStream
{
    void Initialize(IStream pstream, uint grfMode);
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("196bf9a5-b61e-4f7e-b79e-2eb682b29675")]
public interface IPreviewHandlerVisuals
{
    void SetBackgroundColor(uint color);
    void SetFont([In] ref LOGFONT plf);
    void SetTextColor(uint color);
}

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct LOGFONT
{
    public int lfHeight;
    public int lfWidth;
    public int lfEscapement;
    public int lfOrientation;
    public int lfWeight;
    public byte lfItalic;
    public byte lfUnderline;
    public byte lfStrikeOut;
    public byte lfCharSet;
    public byte lfOutPrecision;
    public byte lfClipPrecision;
    public byte lfQuality;
    public byte lfPitchAndFamily;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
    public string lfFaceName;
}

// SHCreateStreamOnFileEx pro IInitializeWithStream
public static class NativeMethods
{
    [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
    public static extern int SHCreateStreamOnFileEx(
        string pszFile, uint grfMode, uint dwAttributes,
        bool fCreate, IntPtr pstmTemplate, out IStream ppstm);

    public const uint STGM_READ = 0x00000000;
    public const uint STGM_SHARE_DENY_NONE = 0x00000040;
}
