# Scratchpad

A lightweight spreadsheet scratchpad for Windows. Quick-launch, small enough to forget it's there, enough formula support to actually use it.

![Scratchpad dark mode](icon-512.png)

## Install

Download `Scratchpad.exe` from the [latest release](https://github.com/CaseyMargell/Scratchpad/releases/latest) and run it. That's it — no installer, no admin, no runtime to install. Self-contained single-file executable (~70MB).

> **Note:** Windows SmartScreen may say "Windows protected your PC" on first run since the exe isn't code-signed. Click **More info** → **Run anyway**. This is normal for unsigned apps.

## Use it

- **Global hotkey** (default `Ctrl+Alt+S`): brings the window to front from anywhere. Press again to hide.
- **Close (X) button** hides to background; the hotkey stays live.
- **Settings → Quit Scratchpad** to fully exit.
- **Start with Windows** option in Settings launches it minimized on sign-in.

### Keyboard

| Key | Action |
| --- | --- |
| `F2` / `Enter` | Edit cell (keeps content selected so typing replaces it) |
| Any printable key | Enter "fresh-type" mode in the selected cell |
| `Tab` / `Shift+Tab` | Next / previous cell |
| `Arrow` keys | Navigate |
| `Shift + Arrow` / `Shift + Click` | Extend selection into a range |
| `Ctrl + Z` / `Ctrl + Y` | Undo / redo (50 steps) |
| `Ctrl + C` / `Ctrl + X` / `Ctrl + V` | Copy / cut / paste (Excel-compatible, tab-delimited) |
| `Ctrl + A` | Select all |
| `Esc` | Cancel current edit |

### Formulas

Type `=` to start a formula. Arrow keys and clicks insert cell references. Examples:

```
=A1+B1                    =SUM(A1:A10)
=IF(A1>100, "big", "ok")  =AVERAGE(B1:B5)
=ROUND(A1*0.08, 2)        =COUNTIF(C1:C10, ">0")
=CONCATENATE(A1, " ", B1) =UPPER(A1)
```

Supported functions: `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTIF`, `SUMIF`, `IF`, `ROUND`, `ABS`, `SQRT`, `POWER`, `MOD`, `CONCATENATE`, `UPPER`, `LOWER`, `TRIM`, `LEN`, `LEFT`, `RIGHT`, `TODAY`, `NOW`. Unclosed parens are auto-closed on Enter.

### Excel-style niceties

- **Drag-to-fill**: hover the selected cell to reveal the blue square, then drag it to copy the formula across a range. Cell references auto-adjust.
- **Text overflow**: long text in one cell spills into empty cells to its right.
- **Auto-size column**: double-click the right edge of a column header to fit its widest content.
- **Drag to resize** columns.
- **Save / Open `.xlsx`** via toolbar icons.

### Settings

- **Theme**: dark or light (icons swap between filled and outline variants).
- **Grid size**: Small (4×6), Medium (8×12), Large (16×24).
- **Global hotkey**: click the current shortcut in Settings and press a new combo to change it.

## Web preview

There's also a pure-HTML version of the earlier implementation at [`scratchpad.html`](scratchpad.html) (root of repo). It lacks the native integration (no global hotkey, no start-with-Windows, etc.) but works on non-Windows platforms and requires no install. Useful for quickly sharing the feel of the app.

## Building from source

Requirements: Windows, .NET 8 SDK.

```powershell
cd app
dotnet publish -c Release
```

Output: `app/bin/Release/net8.0-windows/win-x64/publish/Scratchpad.exe` — self-contained, single-file. Copy anywhere and run.

## Config & data locations

- `%LocalAppData%\Scratchpad\config.json` — hotkey, window bounds, theme, grid size
- Cell data is persisted in the same config file; undo/redo history is in-memory per session.

## License

MIT. See [LICENSE](LICENSE). Fluent UI System Icons used for toolbar icons ([Microsoft, MIT](https://github.com/microsoft/fluentui-system-icons)).
