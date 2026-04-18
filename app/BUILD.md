# Building the native Scratchpad app

## Requirements
- .NET 8 SDK

## Build a single-file self-contained exe

```
cd app
dotnet publish -c Release
```

Output: `app/bin/Release/net8.0-windows/win-x64/publish/Scratchpad.exe`

The resulting exe is self-contained (includes the .NET runtime) and bundles all
HTML/CSS/JS assets as embedded resources. No install step required — copy the
exe anywhere and run.

## Launch modes

- Double-click the exe: window opens, tray icon appears
- `Scratchpad.exe --tray`: starts hidden in tray (used by "Start with Windows")

## First-run notes

- WebView2 Runtime must be installed. Windows 11 ships with it. Windows 10 users
  may need the [Evergreen Bootstrapper](https://developer.microsoft.com/en-us/microsoft-edge/webview2/).
- Windows SmartScreen may warn on an unsigned exe; users can click "More info"
  → "Run anyway".

## Config and data locations

- `%LocalAppData%\Scratchpad\config.json` — hotkey combo, window bounds
- `%LocalAppData%\Scratchpad\WebView2\` — browser profile (localStorage, etc.)
