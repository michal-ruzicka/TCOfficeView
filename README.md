# TC Office Lister Plugin

Total Commander Lister plugin pro náhled MS Office dokumentů (DOCX, XLSX, PPTX...) 
přes nativní **Windows Preview Handlery** registrované samotným Office.

## Architektura

```
┌──────────────────────────────────────────────────────┐
│  Total Commander (F3 / Quick View)                   │
│                                                      │
│  ┌────────────────────────────────────────────┐      │
│  │  tcoffice.wlx(64)   ← TC načítá tuto DLL   │      │
│  │                                            │      │
│  │  • ListLoad()   → vytvoří child HWND       │      │
│  │  • LaunchHost() → spustí host process      │      │
│  │  • Named pipe komunikace                   │      │
│  └────────────────────────────────────────────┘      │
└──────────────────────────────────────────────────────┘
                  │  named pipe
                  ▼
┌──────────────────────────────────────────────────────┐
│  OfficePreviewHost.exe  (izolovaný STA proces)       │
│                                                      │
│  • CoCreateInstance(Word/Excel/PowerPoint           │
│                     Preview Handler CLSID)           │
│  • IInitializeWithFile/Stream::Initialize(path)      │
│  • IPreviewHandler::SetWindow(hwnd) + DoPreview()    │
│                                                      │
│  → Office renderuje přímo do HWND TC                 │
└──────────────────────────────────────────────────────┘
```

## Proč izolovaný proces?

1. **Stabilita** – Office Preview Handlery občas crashují. Crash v host procesu 
   neshodí TC. (Stejný důvod, proč Windows Explorer používá `prevhost.exe`.)
2. **STA threading** – COM preview handlery vyžadují single-threaded apartment, 
   což se obtížně garantuje v rámci pluginu DLL načteného TC.
3. **Bitness flexibilita** – Office handlery jsou registrovány pro konkrétní 
   bitness. 64-bit TC + 32-bit Office? Spustíme 32-bit host process.

## Build

### Předpoklady

- Visual Studio 2022 Build Tools (workload: Desktop development with C++)
- .NET 8 SDK
- CMake 3.20+

### Postup

```cmd
build.cmd
```

Výstupy v `dist\`:
- `tcoffice.wlx` – plugin pro 32-bit TC
- `tcoffice.wlx64` – plugin pro 64-bit TC
- `OfficePreviewHost.exe` – host process (64-bit)
- `OfficePreviewHost_x86.exe` – host process (32-bit)

## Instalace

1. Zkopíruj `dist\*` do nějaké stálé složky, např. `C:\Tools\TCOffice\`.
2. V TC: **Configuration → Options → Plugins → Lister plugins → Configure**.
3. Klikni **Add** a vyber `tcoffice.wlx` (nebo `.wlx64`).
4. TC plugin automaticky zaregistruje pro přípony deklarované v `DetectString`.

Test: F3 na libovolný `.docx`/`.xlsx`/`.pptx` soubor.

## Podporované formáty

Závisí na nainstalovaných Preview Handlerech. S nainstalovaným Office obvykle:

| Aplikace | Přípony |
|----------|---------|
| Word | DOC, DOCX, DOCM, RTF |
| Excel | XLS, XLSX, XLSM, XLSB |
| PowerPoint | PPT, PPTX, PPTM |
| Visio | VSD, VSDX |
| Outlook | MSG |

## Protokol DLL ↔ Host

Textový, UTF-16 LE, řádky zakončené `\n`.

**Plugin → Host:**
```
LOAD C:\path\to\file.docx
RESIZE 1024 768
CLOSE
```

**Host → Plugin:**
```
OK
ERR <message>
```

## Známé limitace

- **Performance první spuštění:** První preview po startu TC trvá ~500-1500 ms 
  (cold start Office COM serveru). Následné jsou rychlé (~100-300 ms).
- **Memory:** Každý host process drží ~40-100 MB. Při procházení mnoha souborů 
  je možné implementovat reuse host procesu napříč `ListLoadNext` voláními 
  (už implementováno – `ListLoadNextW` pošle nový LOAD do běžícího hostu).
- **Print/Search/Copy:** Zatím neimplementováno (`ListSendCommand` returns 0). 
  Office Preview Handlery samy poskytují vlastní kontextové menu uvnitř.

## Troubleshooting

**Plugin se neaktivuje:** Zkontroluj v TC Lister Plugin dialog, jestli je plugin 
nahrán a přípona je v "Detect string".

**Plugin se aktivuje ale okno je prázdné:** Pravděpodobně chybí Preview Handler 
pro daný typ. Ověř v registru: 
`HKCR\.docx\shellex\{8895b1c6-b41f-4c1c-a562-0d564250836f}`

**Crash TC:** Nemělo by se stát díky izolaci hostu, ale pokud ano, podívej se 
do Event Vieweru na crash `OfficePreviewHost.exe`. Často je důvodem 64/32-bit 
mismatch (např. 64-bit Office, ale plugin spustil 32-bit hosta).
