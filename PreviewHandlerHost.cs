using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace TCOffice.Host;

/// <summary>
/// Wrapper kolem IPreviewHandler. Najde správný handler pro daný typ souboru,
/// inicializuje ho a hostuje v zadaném HWND.
/// </summary>
public sealed class PreviewHandlerHost : IDisposable
{
    // GUID kategorie Preview Handlerů v registru
    private static readonly Guid PreviewHandlerCategory =
        new("8895b1c6-b41f-4c1c-a562-0d564250836f");

    private IPreviewHandler? _handler;
    private object? _comObject;
    private readonly IntPtr _hwnd;

    public PreviewHandlerHost(IntPtr hwnd) { _hwnd = hwnd; }

    /// <summary>
    /// Najde CLSID Preview Handleru pro danou příponu v registru.
    /// Hledá v HKCU i HKLM, fallback přes HKCR\.{ext}\PerceivedType.
    /// </summary>
    public static Guid? FindHandlerClsid(string filePath)
    {
        string ext = Path.GetExtension(filePath).ToLowerInvariant();
        if (string.IsNullOrEmpty(ext)) return null;

        // Cesta: HKCR\.{ext}\shellex\{PreviewHandlerCategory}
        string subPath = $@"{ext}\shellex\{{{PreviewHandlerCategory:D}}}";

        foreach (var root in new[] { Registry.ClassesRoot,
                                     Registry.CurrentUser.OpenSubKey(@"Software\Classes"),
                                     Registry.LocalMachine.OpenSubKey(@"Software\Classes") })
        {
            if (root == null) continue;
            using var key = root.OpenSubKey(subPath);
            string? value = key?.GetValue(null) as string;
            if (Guid.TryParse(value, out var guid)) return guid;
        }

        // Některé typy (např. .docx) mají handler navázaný přes ProgID
        using var extKey = Registry.ClassesRoot.OpenSubKey(ext);
        string? progId = extKey?.GetValue(null) as string;
        if (!string.IsNullOrEmpty(progId))
        {
            using var progKey = Registry.ClassesRoot.OpenSubKey(
                $@"{progId}\shellex\{{{PreviewHandlerCategory:D}}}");
            string? value = progKey?.GetValue(null) as string;
            if (Guid.TryParse(value, out var guid)) return guid;
        }

        return null;
    }

    /// <summary>
    /// Načte daný soubor do preview handleru a vyrenderuje ho.
    /// </summary>
    public void LoadFile(string filePath)
    {
        Unload();

        Guid clsid = FindHandlerClsid(filePath)
            ?? throw new InvalidOperationException(
                $"Pro soubor '{Path.GetFileName(filePath)}' není zaregistrován Preview Handler.");

        // Vytvoř COM objekt
        Type? type = Type.GetTypeFromCLSID(clsid)
            ?? throw new InvalidOperationException($"CLSID {clsid} nelze vytvořit.");

        _comObject = Activator.CreateInstance(type)
            ?? throw new InvalidOperationException("CoCreateInstance vrátil null.");

        // Inicializace - preview handlery podporují různá rozhraní podle typu
        if (_comObject is IInitializeWithFile fileInit)
        {
            fileInit.Initialize(filePath, NativeMethods.STGM_READ);
        }
        else if (_comObject is IInitializeWithStream streamInit)
        {
            int hr = NativeMethods.SHCreateStreamOnFileEx(
                filePath,
                NativeMethods.STGM_READ | NativeMethods.STGM_SHARE_DENY_NONE,
                0, false, IntPtr.Zero, out var stream);
            Marshal.ThrowExceptionForHR(hr);
            streamInit.Initialize(stream, NativeMethods.STGM_READ);
        }
        else
        {
            throw new InvalidOperationException(
                "Preview handler nepodporuje IInitializeWithFile ani IInitializeWithStream.");
        }

        _handler = (IPreviewHandler)_comObject;

        // Nastav okno a vyrenderuj
        var rect = GetClientRect(_hwnd);
        _handler.SetWindow(_hwnd, ref rect);
        _handler.DoPreview();
    }

    public void Resize(int width, int height)
    {
        if (_handler == null) return;
        var rect = new RECT(0, 0, width, height);
        _handler.SetRect(ref rect);
    }

    public void Unload()
    {
        if (_handler != null)
        {
            try { _handler.Unload(); } catch { /* handler může být ve špatném stavu */ }
            _handler = null;
        }
        if (_comObject != null)
        {
            Marshal.FinalReleaseComObject(_comObject);
            _comObject = null;
        }
    }

    public void Dispose() => Unload();

    [DllImport("user32.dll")]
    private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

    private static RECT GetClientRect(IntPtr hwnd)
    {
        GetClientRect(hwnd, out var r);
        return r;
    }
}
