using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Windows;
using System.Windows.Interop;
using Microsoft.Web.WebView2.Core;

namespace Scratchpad;

public partial class MainWindow : Window
{
    private HotkeyManager? _hotkey;
    private HwndSource? _hwndSource;
    private bool _realClose;

    public MainWindow()
    {
        InitializeComponent();
        Loaded += MainWindow_Loaded;
        Closing += MainWindow_Closing;
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        RestoreWindowBounds();

        // Use a user data folder next to the exe so settings persist
        var userDataDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Scratchpad", "WebView2");
        Directory.CreateDirectory(userDataDir);

        var env = await CoreWebView2Environment.CreateAsync(null, userDataDir);
        await WebView.EnsureCoreWebView2Async(env);

        var core = WebView.CoreWebView2;

        // Extract embedded HTML resources to a temp folder and serve via virtual host
        var webRoot = ExtractWebAssets();
        core.SetVirtualHostNameToFolderMapping("scratchpad.local", webRoot,
            CoreWebView2HostResourceAccessKind.Allow);

        // JS -> C# bridge
        core.WebMessageReceived += OnWebMessageReceived;

        // Inject a tiny JS bridge API as soon as the page loads
        await core.AddScriptToExecuteOnDocumentCreatedAsync(@"
            window.NativeBridge = {
                send: (msg) => window.chrome.webview.postMessage(JSON.stringify(msg)),
                isNative: true
            };
            window.addEventListener('DOMContentLoaded', () => {
                document.body.classList.add('native-app');
            });
        ");

        core.Navigate("https://scratchpad.local/scratchpad.html");

        // Hotkey registration happens once HWND is ready
        var helper = new WindowInteropHelper(this);
        _hwndSource = HwndSource.FromHwnd(helper.Handle);
        _hwndSource?.AddHook(WndProc);

        _hotkey = new HotkeyManager(helper.Handle);
        var combo = Settings.LoadHotkey();
        _hotkey.Register(combo.Modifiers, combo.Key);
    }

    private string ExtractWebAssets()
    {
        var webRoot = Path.Combine(Path.GetTempPath(), "Scratchpad-web");
        Directory.CreateDirectory(webRoot);

        var asm = Assembly.GetExecutingAssembly();
        string[] files = { "scratchpad.html", "manifest.json", "icon-192.png", "icon-512.png", "icon.svg" };
        foreach (var f in files)
        {
            using var s = asm.GetManifestResourceStream(f);
            if (s == null) continue;
            var dest = Path.Combine(webRoot, f);
            using var outFs = File.Create(dest);
            s.CopyTo(outFs);
        }
        return webRoot;
    }

    private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
    {
        try
        {
            using var doc = JsonDocument.Parse(e.WebMessageAsJson);
            var root = doc.RootElement;
            // WebView wraps JS strings in quotes — parse inner JSON
            if (root.ValueKind == JsonValueKind.String)
            {
                using var inner = JsonDocument.Parse(root.GetString()!);
                HandleMessage(inner.RootElement);
            }
            else
            {
                HandleMessage(root);
            }
        }
        catch
        {
            // Ignore malformed messages
        }
    }

    private void HandleMessage(JsonElement msg)
    {
        if (!msg.TryGetProperty("type", out var typeEl)) return;
        var type = typeEl.GetString();

        switch (type)
        {
            case "setHotkey":
                if (msg.TryGetProperty("modifiers", out var modEl) &&
                    msg.TryGetProperty("key", out var keyEl))
                {
                    var mods = (uint)modEl.GetInt32();
                    var key = (uint)keyEl.GetInt32();
                    _hotkey?.Unregister();
                    var success = _hotkey?.Register(mods, key) ?? false;
                    if (success) Settings.SaveHotkey(mods, key);
                    SendToJs(new { type = "hotkeyResult", success });
                }
                break;

            case "setStartup":
                if (msg.TryGetProperty("enabled", out var enEl))
                {
                    Settings.SetRunAtStartup(enEl.GetBoolean());
                }
                break;

            case "getSettings":
                var hk = Settings.LoadHotkey();
                SendToJs(new
                {
                    type = "settings",
                    hotkey = new { modifiers = hk.Modifiers, key = hk.Key },
                    runAtStartup = Settings.GetRunAtStartup(),
                    version = GetVersion()
                });
                break;

            case "minimizeToTray":
                Hide();
                break;

            case "resizeContent":
                if (msg.TryGetProperty("width", out var wEl) &&
                    msg.TryGetProperty("height", out var hEl))
                {
                    ResizeToContent(wEl.GetDouble(), hEl.GetDouble());
                }
                break;

            case "resizeForGridIfUnsized":
                if (Settings.LoadWindowBounds() == null &&
                    msg.TryGetProperty("width", out var iw) &&
                    msg.TryGetProperty("height", out var ih))
                {
                    ResizeToContent(iw.GetDouble(), ih.GetDouble());
                }
                break;
        }
    }

    private void ResizeToContent(double contentW, double contentH)
    {
        var deltaW = ActualWidth - WebView.ActualWidth;
        var deltaH = ActualHeight - WebView.ActualHeight;
        if (deltaW < 0 || deltaH < 0) { deltaW = 16; deltaH = 40; }
        Width = Math.Max(MinWidth, contentW + deltaW);
        Height = Math.Max(MinHeight, contentH + deltaH);
        SaveWindowBounds();
    }

    private void SendToJs(object payload)
    {
        var json = JsonSerializer.Serialize(payload);
        WebView.CoreWebView2?.PostWebMessageAsJson(json);
    }

    private static string GetVersion()
    {
        var v = Assembly.GetExecutingAssembly().GetName().Version;
        return v == null ? "dev" : $"{v.Major}.{v.Minor}.{v.Build}";
    }

    private const int WM_HOTKEY = 0x0312;

    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WM_HOTKEY)
        {
            ToggleVisibility();
            handled = true;
        }
        return IntPtr.Zero;
    }

    public void ToggleVisibility()
    {
        if (IsVisible && WindowState != WindowState.Minimized && IsActive)
        {
            Hide();
        }
        else
        {
            Show();
            if (WindowState == WindowState.Minimized) WindowState = WindowState.Normal;
            Activate();
            Topmost = true;
            Topmost = false;
            Focus();
        }
    }

    private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        if (_realClose) return;
        e.Cancel = true;
        SaveWindowBounds();
        Hide();
    }

    public void ReallyClose()
    {
        _realClose = true;
        SaveWindowBounds();
        _hotkey?.Dispose();
        Close();
        Application.Current.Shutdown();
    }

    private void RestoreWindowBounds()
    {
        var b = Settings.LoadWindowBounds();
        if (b.HasValue)
        {
            var (l, t, w, h) = b.Value;
            Left = l; Top = t; Width = w; Height = h;
            WindowStartupLocation = WindowStartupLocation.Manual;
        }
    }

    private void SaveWindowBounds()
    {
        if (WindowState == WindowState.Normal)
            Settings.SaveWindowBounds(Left, Top, Width, Height);
    }
}
