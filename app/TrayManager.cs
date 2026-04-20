using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using H.NotifyIcon;

namespace Scratchpad;

public class TrayManager : IDisposable
{
    private readonly TaskbarIcon _tray;
    private readonly MainWindow _main;

    public TrayManager(MainWindow main)
    {
        _main = main;

        _tray = new TaskbarIcon
        {
            ToolTipText = "Scratchpad",
            Icon = LoadTrayIcon()
        };

        var menu = new ContextMenu();
        var showItem = new MenuItem { Header = "Show" };
        showItem.Click += (_, _) => _main.ToggleVisibility();
        var quitItem = new MenuItem { Header = "Quit" };
        quitItem.Click += (_, _) => _main.ReallyClose();
        menu.Items.Add(showItem);
        menu.Items.Add(new Separator());
        menu.Items.Add(quitItem);

        _tray.ContextMenu = menu;
        _tray.TrayLeftMouseDown += (_, _) => _main.ToggleVisibility();
    }

    private Icon LoadTrayIcon()
    {
        // Prefer the exe's embedded app icon (app.ico, multi-res, crisp at 16x16)
        try
        {
            var exe = Environment.ProcessPath;
            if (!string.IsNullOrEmpty(exe))
            {
                var icon = Icon.ExtractAssociatedIcon(exe);
                if (icon != null) return icon;
            }
        }
        catch { }

        // Fallback: convert the embedded 192px PNG
        try
        {
            var asm = Assembly.GetExecutingAssembly();
            using var stream = asm.GetManifestResourceStream("icon-192.png");
            if (stream != null)
            {
                using var bmp = new Bitmap(stream);
                var hIcon = bmp.GetHicon();
                return Icon.FromHandle(hIcon);
            }
        }
        catch { }

        return SystemIcons.Application;
    }

    public void Dispose()
    {
        _tray.Dispose();
    }
}
