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
    private bool _firstShown;
    private bool _quitting;

    public bool StartMinimized { get; set; }

    public MainWindow()
    {
        InitializeComponent();
        // Start invisible — the window flashes black→white during WebView2 init
        // if we show before content is ready. We'll call Show() once the HTML renders.
        Visibility = Visibility.Hidden;
        WindowStartupLocation = WindowStartupLocation.Manual;
        Left = -32000; Top = -32000;  // off-screen until we decide where to put it
        SourceInitialized += MainWindow_SourceInitialized;
        Loaded += MainWindow_Loaded;
        Closing += MainWindow_Closing;
    }

    private void MainWindow_SourceInitialized(object? sender, EventArgs e)
    {
        // Register global hotkey as early as possible so it works even before window is shown
        var helper = new WindowInteropHelper(this);
        _hwndSource = HwndSource.FromHwnd(helper.Handle);
        _hwndSource?.AddHook(WndProc);

        _hotkey = new HotkeyManager(helper.Handle);
        var combo = Settings.LoadHotkey();
        _hotkey.Register(combo.Modifiers, combo.Key);
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        var userDataDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Scratchpad", "WebView2");
        Directory.CreateDirectory(userDataDir);

        var env = await CoreWebView2Environment.CreateAsync(null, userDataDir);
        await WebView.EnsureCoreWebView2Async(env);

        var core = WebView.CoreWebView2;

        var webRoot = ExtractWebAssets();
        core.SetVirtualHostNameToFolderMapping("scratchpad.local", webRoot,
            CoreWebView2HostResourceAccessKind.Allow);

        core.WebMessageReceived += OnWebMessageReceived;

        await core.AddScriptToExecuteOnDocumentCreatedAsync(@"
            window.NativeBridge = {
                send: (msg) => window.chrome.webview.postMessage(JSON.stringify(msg)),
                isNative: true
            };
            window.addEventListener('DOMContentLoaded', () => {
                document.body.classList.add('native-app');
            });
        ");

        // Show window only after the DOM is ready, so there's no flash
        core.DOMContentLoaded += (_, __) =>
        {
            if (_firstShown) return;
            _firstShown = true;
            Dispatcher.InvokeAsync(ShowFirstTime);
        };

        core.Navigate("https://scratchpad.local/scratchpad.html");
    }

    private void ShowFirstTime()
    {
        RestoreWindowBounds();
        Visibility = Visibility.Visible;
        if (StartMinimized)
        {
            WindowState = WindowState.Minimized;
        }
        else
        {
            WindowState = WindowState.Normal;
            Activate();
        }
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

            case "hideWindow":
                Hide();
                break;

            case "quitApp":
                Quit();
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
            if (IsVisible && IsActive && WindowState != WindowState.Minimized)
            {
                Hide();
            }
            else
            {
                BringToFront();
            }
            handled = true;
        }
        return IntPtr.Zero;
    }

    public void BringToFront()
    {
        Show();
        if (WindowState == WindowState.Minimized) WindowState = WindowState.Normal;
        Activate();
        Topmost = true;
        Topmost = false;
        Focus();
    }

    private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        if (_quitting) return;
        // X button hides the window; process keeps running so the hotkey stays warm
        e.Cancel = true;
        SaveWindowBounds();
        Hide();
    }

    public void Quit()
    {
        _quitting = true;
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
        }
        else
        {
            // Center on primary screen if no saved bounds
            Left = (SystemParameters.WorkArea.Width - Width) / 2;
            Top = (SystemParameters.WorkArea.Height - Height) / 2;
        }
    }

    private void SaveWindowBounds()
    {
        if (WindowState == WindowState.Normal && _firstShown)
            Settings.SaveWindowBounds(Left, Top, Width, Height);
    }
}
