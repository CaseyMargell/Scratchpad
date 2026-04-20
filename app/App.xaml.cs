using System.IO;
using System.IO.Pipes;
using System.Windows;

namespace Scratchpad;

public partial class App : Application
{
    public static MainWindow? MainWin { get; private set; }

    private const string MutexName = "Scratchpad.SingleInstance.Mutex.v1";
    private const string PipeName = "Scratchpad.SingleInstance.Pipe.v1";
    private static System.Threading.Mutex? _mutex;

    private void App_Startup(object sender, StartupEventArgs e)
    {
        bool startMinimized = e.Args.Any(a => a.Equals("--minimized", StringComparison.OrdinalIgnoreCase));

        // Single instance check: if already running, ping it and exit
        _mutex = new System.Threading.Mutex(true, MutexName, out var isFirstInstance);
        if (!isFirstInstance)
        {
            try
            {
                using var client = new NamedPipeClientStream(".", PipeName, PipeDirection.Out);
                client.Connect(500);
                using var writer = new StreamWriter(client);
                writer.WriteLine("show");
                writer.Flush();
            }
            catch { }
            Shutdown();
            return;
        }

        // Start listening for activation signals from other launches
        Task.Run(ListenForSecondInstance);

        MainWin = new MainWindow();
        MainWin.StartMinimized = startMinimized;
        // Show the window off-screen so the HWND exists (needed for hotkey registration
        // and taskbar entry). MainWindow will move it on-screen once WebView2 is ready.
        MainWin.Show();
    }

    private async Task ListenForSecondInstance()
    {
        while (true)
        {
            try
            {
                using var server = new NamedPipeServerStream(PipeName, PipeDirection.In);
                await server.WaitForConnectionAsync();
                using var reader = new StreamReader(server);
                var msg = await reader.ReadLineAsync();
                if (msg == "show")
                {
                    Dispatcher.Invoke(() => MainWin?.BringToFront());
                }
            }
            catch { }
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _mutex?.ReleaseMutex();
        _mutex?.Dispose();
        base.OnExit(e);
    }
}
