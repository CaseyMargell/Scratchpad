using System.Reflection;
using System.Windows;
using System.Windows.Input;

namespace Scratchpad;

public partial class SettingsWindow : Window
{
    private readonly MainWindow _main;
    private bool _capturing;
    private HotkeyCombo _pendingHotkey;

    public SettingsWindow(MainWindow main)
    {
        InitializeComponent();
        _main = main;
        var s = AppSettings.Load();
        _pendingHotkey = s.Hotkey;

        SetActive(DarkBtn, !s.IsLight);
        SetActive(LightBtn, s.IsLight);

        SetActive(SmallBtn,  s.GridSize == GridSize.Small);
        SetActive(MediumBtn, s.GridSize == GridSize.Medium);
        SetActive(LargeBtn,  s.GridSize == GridSize.Large);

        HotkeyBtn.Content = FormatHotkey(s.Hotkey);

        var startup = AppSettings.GetRunAtStartup();
        SetActive(StartupOffBtn, !startup);
        SetActive(StartupOnBtn, startup);

        var v = Assembly.GetExecutingAssembly().GetName().Version;
        VersionText.Text = v == null ? "dev" : $"v{v.Major}.{v.Minor}.{v.Build}";

        PreviewKeyDown += SettingsWindow_PreviewKeyDown;
    }

    private static void SetActive(System.Windows.Controls.Button b, bool active)
    {
        b.Tag = active ? "active" : null;
    }

    private void DarkBtn_Click(object sender, RoutedEventArgs e) { _main.ApplyTheme(false); SetActive(DarkBtn, true); SetActive(LightBtn, false); }
    private void LightBtn_Click(object sender, RoutedEventArgs e) { _main.ApplyTheme(true); SetActive(DarkBtn, false); SetActive(LightBtn, true); }

    private void SizeBtn_Click(object sender, RoutedEventArgs e)
    {
        var s = sender == SmallBtn ? GridSize.Small : sender == MediumBtn ? GridSize.Medium : GridSize.Large;
        _main.ApplyGridSize(s);
        // Reload the current actual size (may have been cancelled by user)
        var now = AppSettings.Load().GridSize;
        SetActive(SmallBtn,  now == GridSize.Small);
        SetActive(MediumBtn, now == GridSize.Medium);
        SetActive(LargeBtn,  now == GridSize.Large);
    }

    private void HotkeyBtn_Click(object sender, RoutedEventArgs e)
    {
        _capturing = true;
        HotkeyBtn.Content = "Press new shortcut…";
    }

    private void SettingsWindow_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (!_capturing) return;
        if (e.Key == Key.Escape)
        {
            _capturing = false;
            HotkeyBtn.Content = FormatHotkey(_pendingHotkey);
            e.Handled = true;
            return;
        }
        var mods = Keyboard.Modifiers;
        if (!mods.HasFlag(ModifierKeys.Control) && !mods.HasFlag(ModifierKeys.Alt) && !mods.HasFlag(ModifierKeys.Windows)) return;
        if (e.Key == Key.LeftCtrl || e.Key == Key.RightCtrl ||
            e.Key == Key.LeftAlt || e.Key == Key.RightAlt ||
            e.Key == Key.LeftShift || e.Key == Key.RightShift ||
            e.Key == Key.LWin || e.Key == Key.RWin) return;

        uint m = 0;
        if (mods.HasFlag(ModifierKeys.Alt)) m |= HotkeyManager.MOD_ALT;
        if (mods.HasFlag(ModifierKeys.Control)) m |= HotkeyManager.MOD_CONTROL;
        if (mods.HasFlag(ModifierKeys.Shift)) m |= HotkeyManager.MOD_SHIFT;
        if (mods.HasFlag(ModifierKeys.Windows)) m |= HotkeyManager.MOD_WIN;
        var vk = (uint)KeyInterop.VirtualKeyFromKey(e.Key);

        var newCombo = new HotkeyCombo(m, vk);
        // Try to register via the main window
        var s = AppSettings.Load();
        var oldCombo = s.Hotkey;
        s.Hotkey = newCombo;
        AppSettings.Save(s);
        // Ping main to re-register
        if (!_main.TryApplyHotkey(newCombo))
        {
            // Restore old
            s.Hotkey = oldCombo;
            AppSettings.Save(s);
            HotkeyBtn.Content = "Conflict — try another";
            _capturing = false;
            e.Handled = true;
            return;
        }
        _pendingHotkey = newCombo;
        HotkeyBtn.Content = FormatHotkey(newCombo);
        _capturing = false;
        e.Handled = true;
    }

    private void StartupOffBtn_Click(object sender, RoutedEventArgs e)
    {
        AppSettings.SetRunAtStartup(false);
        SetActive(StartupOffBtn, true);
        SetActive(StartupOnBtn, false);
    }

    private void StartupOnBtn_Click(object sender, RoutedEventArgs e)
    {
        AppSettings.SetRunAtStartup(true);
        SetActive(StartupOffBtn, false);
        SetActive(StartupOnBtn, true);
    }

    private void QuitBtn_Click(object sender, RoutedEventArgs e)
    {
        if (MessageBox.Show(this,
            "Quit Scratchpad?\n\nThe global hotkey will stop working until you relaunch.",
            "Confirm quit", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
        Close();
        _main.Quit();
    }

    private static string FormatHotkey(HotkeyCombo c)
    {
        var parts = new List<string>();
        if ((c.Modifiers & HotkeyManager.MOD_CONTROL) != 0) parts.Add("Ctrl");
        if ((c.Modifiers & HotkeyManager.MOD_ALT) != 0) parts.Add("Alt");
        if ((c.Modifiers & HotkeyManager.MOD_SHIFT) != 0) parts.Add("Shift");
        if ((c.Modifiers & HotkeyManager.MOD_WIN) != 0) parts.Add("Win");
        parts.Add(KeyLabel(c.Key));
        return string.Join(" + ", parts);
    }

    private static string KeyLabel(uint vk)
    {
        if (vk >= 0x30 && vk <= 0x39) return ((char)vk).ToString();
        if (vk >= 0x41 && vk <= 0x5A) return ((char)vk).ToString();
        if (vk >= 0x70 && vk <= 0x7B) return "F" + (vk - 0x6F);
        return vk switch
        {
            0x20 => "Space", 0x0D => "Enter", 0x09 => "Tab",
            0xC0 => "`", 0xBD => "-", 0xBB => "=",
            0xDB => "[", 0xDD => "]", 0xDC => "\\",
            0xBA => ";", 0xDE => "'", 0xBC => ",", 0xBE => ".", 0xBF => "/",
            _ => "Key"
        };
    }
}
