using System.Windows;

namespace Scratchpad;

public partial class App : Application
{
    public static MainWindow? MainWin { get; private set; }
    public static TrayManager? Tray { get; private set; }

    private void App_Startup(object sender, StartupEventArgs e)
    {
        bool startMinimized = e.Args.Any(a => a.Equals("--tray", StringComparison.OrdinalIgnoreCase));

        MainWin = new MainWindow();
        Tray = new TrayManager(MainWin);

        if (startMinimized)
        {
            MainWin.Hide();
        }
        else
        {
            MainWin.Show();
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        Tray?.Dispose();
        base.OnExit(e);
    }
}
