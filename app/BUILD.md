# Building the native Scratchpad app

## Requirements
- .NET 8 SDK

## Build a single-file self-contained exe

```
cd app
dotnet publish -c Release
```

Output: `app/bin/Release/net8.0-windows/win-x64/publish/Scratchpad.exe`

Everything (runtime, WPF, ClosedXML, icons) is bundled into one file.
Copy anywhere and run.

## Launch modes

- Double-click the exe: window opens at last saved bounds (or centered on first run)
- `Scratchpad.exe --minimized`: starts minimized (used by "Start with Windows")

## Config & data

- `%LocalAppData%\Scratchpad\config.json` — theme, hotkey, grid size, window bounds
- No data persists outside what you save to .xlsx — this is a scratchpad, after all
