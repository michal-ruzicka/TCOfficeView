using System;
using System.Globalization;
using System.IO;
using System.IO.Pipes;
using System.Threading;

namespace TCOffice.Host;

/// <summary>
/// OfficePreviewHost.exe
///
/// Spouštěn z tcoffice.wlx pluginu. Hostuje Office Preview Handler v zadaném
/// HWND a komunikuje s pluginem přes named pipe.
///
/// Protokol (textový, řádky ukončené \n, kódování UTF-16):
///   LOAD &lt;path&gt;        - načte soubor a vyrenderuje preview
///   RESIZE &lt;w&gt; &lt;h&gt;     - změna velikosti okna (předáno z WM_SIZE)
///   CLOSE              - čistý exit
///
/// Odpovědi:
///   OK
///   ERR &lt;message&gt;
/// </summary>
public static class Program
{
    [STAThread]
    public static int Main(string[] args)
    {
        try
        {
            var (hwnd, pipeName) = ParseArgs(args);
            RunMessageLoop(hwnd, pipeName);
            return 0;
        }
        catch (Exception ex)
        {
            // Fatal - zapiš do event log nebo stderr (TC to neuvidí, ale debug)
            Console.Error.WriteLine($"FATAL: {ex}");
            return 1;
        }
    }

    private static (IntPtr hwnd, string pipeName) ParseArgs(string[] args)
    {
        IntPtr hwnd = IntPtr.Zero;
        string pipeName = "";
        for (int i = 0; i < args.Length - 1; i++)
        {
            switch (args[i])
            {
                case "--hwnd":
                    hwnd = new IntPtr(long.Parse(args[i + 1], CultureInfo.InvariantCulture));
                    break;
                case "--pipe":
                    pipeName = args[i + 1];
                    break;
            }
        }
        if (hwnd == IntPtr.Zero || string.IsNullOrEmpty(pipeName))
            throw new ArgumentException("Usage: --hwnd <handle> --pipe <name>");

        // Plugin nám předal plnou cestu \\.\pipe\name - pro NamedPipeClient
        // chceme jen jméno bez prefixu
        if (pipeName.StartsWith(@"\\.\pipe\"))
            pipeName = pipeName.Substring(@"\\.\pipe\".Length);

        return (hwnd, pipeName);
    }

    private static void RunMessageLoop(IntPtr hwnd, string pipeName)
    {
        using var pipe = new NamedPipeClientStream(
            ".", pipeName, PipeDirection.InOut, PipeOptions.None);
        pipe.Connect(5000);
        pipe.ReadMode = PipeTransmissionMode.Message;

        using var host = new PreviewHandlerHost(hwnd);
        using var reader = new StreamReader(pipe, System.Text.Encoding.Unicode);
        using var writer = new StreamWriter(pipe, System.Text.Encoding.Unicode) { AutoFlush = true };

        string? line;
        while ((line = reader.ReadLine()) != null)
        {
            try
            {
                if (line.StartsWith("LOAD ", StringComparison.Ordinal))
                {
                    string path = line.Substring(5);
                    host.LoadFile(path);
                    writer.WriteLine("OK");
                }
                else if (line.StartsWith("RESIZE ", StringComparison.Ordinal))
                {
                    var parts = line.Substring(7).Split(' ');
                    int w = int.Parse(parts[0], CultureInfo.InvariantCulture);
                    int h = int.Parse(parts[1], CultureInfo.InvariantCulture);
                    host.Resize(w, h);
                    // Resize neodpovídáme - vysoká frekvence při tažení rámu
                }
                else if (line.StartsWith("CLOSE", StringComparison.Ordinal))
                {
                    host.Unload();
                    writer.WriteLine("OK");
                    break;
                }
            }
            catch (Exception ex)
            {
                try { writer.WriteLine($"ERR {ex.Message}"); } catch { }
            }
        }
    }
}
