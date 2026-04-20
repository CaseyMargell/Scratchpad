using System.Windows;

namespace Scratchpad;

public partial class App : Application
{
    public static MainWindow? MainWin { get; private set; }

    private void App_Startup(object sender, StartupEventArgs e)
    {
        bool startMinimized = e.Args.Any(a => a.Equals("--minimized", StringComparison.OrdinalIgnoreCase));

        MainWin = new MainWindow();
        MainWin.Show();
        if (startMinimized) MainWin.WindowState = WindowState.Minimized;
    }
}
