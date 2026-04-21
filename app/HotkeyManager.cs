using System.Runtime.InteropServices;

namespace Scratchpad;

public class HotkeyManager : IDisposable
{
    public const uint MOD_ALT     = 0x0001;
    public const uint MOD_CONTROL = 0x0002;
    public const uint MOD_SHIFT   = 0x0004;
    public const uint MOD_WIN     = 0x0008;
    public const uint MOD_NOREPEAT = 0x4000;

    private const int HOTKEY_ID = 0x1338;

    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    private readonly IntPtr _hwnd;
    private bool _registered;

    public HotkeyManager(IntPtr hwnd) { _hwnd = hwnd; }

    public bool Register(uint modifiers, uint virtualKey)
    {
        Unregister();
        _registered = RegisterHotKey(_hwnd, HOTKEY_ID, modifiers | MOD_NOREPEAT, virtualKey);
        return _registered;
    }

    public void Unregister()
    {
        if (_registered) { UnregisterHotKey(_hwnd, HOTKEY_ID); _registered = false; }
    }

    public void Dispose() => Unregister();
}
