using System.IO;
using System.Text.Json;
using Microsoft.Win32;

namespace Scratchpad;

public record HotkeyCombo(uint Modifiers, uint Key);

public static class Settings
{
    private static readonly string ConfigPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Scratchpad", "config.json");

    private const string StartupRegPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string StartupRegName = "Scratchpad";

    private record ConfigData(
        uint HotkeyModifiers,
        uint HotkeyKey,
        double? WindowLeft,
        double? WindowTop,
        double? WindowWidth,
        double? WindowHeight
    );

    private static ConfigData Load()
    {
        try
        {
            if (File.Exists(ConfigPath))
            {
                var json = File.ReadAllText(ConfigPath);
                var c = JsonSerializer.Deserialize<ConfigData>(json);
                if (c != null) return c;
            }
        }
        catch { }
        // Default: Ctrl+Alt+S (VK_S = 0x53)
        return new ConfigData(
            HotkeyManager.MOD_CONTROL | HotkeyManager.MOD_ALT,
            0x53, null, null, null, null);
    }

    private static void Save(ConfigData c)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(ConfigPath)!);
            File.WriteAllText(ConfigPath, JsonSerializer.Serialize(c));
        }
        catch { }
    }

    public static HotkeyCombo LoadHotkey()
    {
        var c = Load();
        return new HotkeyCombo(c.HotkeyModifiers, c.HotkeyKey);
    }

    public static void SaveHotkey(uint modifiers, uint key)
    {
        var c = Load();
        Save(c with { HotkeyModifiers = modifiers, HotkeyKey = key });
    }

    public static (double left, double top, double width, double height)? LoadWindowBounds()
    {
        var c = Load();
        if (c.WindowLeft.HasValue && c.WindowTop.HasValue && c.WindowWidth.HasValue && c.WindowHeight.HasValue)
            return (c.WindowLeft.Value, c.WindowTop.Value, c.WindowWidth.Value, c.WindowHeight.Value);
        return null;
    }

    public static void SaveWindowBounds(double left, double top, double width, double height)
    {
        var c = Load();
        Save(c with { WindowLeft = left, WindowTop = top, WindowWidth = width, WindowHeight = height });
    }

    public static bool GetRunAtStartup()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(StartupRegPath);
            return key?.GetValue(StartupRegName) != null;
        }
        catch { return false; }
    }

    public static void SetRunAtStartup(bool enabled)
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(StartupRegPath, writable: true);
            if (key == null) return;
            if (enabled)
            {
                var exe = Environment.ProcessPath ?? "";
                key.SetValue(StartupRegName, $"\"{exe}\" --minimized");
            }
            else
            {
                key.DeleteValue(StartupRegName, throwOnMissingValue: false);
            }
        }
        catch { }
    }
}
