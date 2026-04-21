using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows;
using Microsoft.Win32;

namespace Scratchpad;

public class AppSettings
{
    public GridSize GridSize { get; set; } = GridSize.Medium;
    public bool IsLight { get; set; }
    public HotkeyCombo Hotkey { get; set; } = new(HotkeyManager.MOD_CONTROL | HotkeyManager.MOD_ALT, 0x53);
    public Rect? WindowBounds { get; set; }
    public string? LastSelectedCell { get; set; }
    public string? LastOpenedFile { get; set; }

    private static readonly string ConfigPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Scratchpad", "config.json");

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() }
    };

    private static AppSettings? _cached;

    public static AppSettings Load()
    {
        if (_cached != null) return _cached;
        try
        {
            if (File.Exists(ConfigPath))
            {
                var json = File.ReadAllText(ConfigPath);
                _cached = JsonSerializer.Deserialize<AppSettings>(json, JsonOpts) ?? new();
                return _cached;
            }
        }
        catch { }
        _cached = new AppSettings();
        return _cached;
    }

    public static void Save(AppSettings s)
    {
        _cached = s;
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(ConfigPath)!);
            File.WriteAllText(ConfigPath, JsonSerializer.Serialize(s, JsonOpts));
        }
        catch { }
    }

    // --- Run at startup ---
    private const string StartupRegPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
    private const string StartupRegName = "Scratchpad";

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

public record HotkeyCombo(uint Modifiers, uint Key);
