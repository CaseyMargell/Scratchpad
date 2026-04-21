using System.IO;
using System.IO.Pipes;
using System.Windows;

namespace Scratchpad;

public partial class App : Application
{
    public static MainWindow? MainWin { get; private set; }

    private const string MutexName = "Scratchpad.SingleInstance.Mutex.v2";
    private const string PipeName = "Scratchpad.SingleInstance.Pipe.v2";
    private static System.Threading.Mutex? _mutex;
    private static bool _ownsMutex;

    private void App_Startup(object sender, StartupEventArgs e)
    {
        bool startMinimized = e.Args.Any(a => a.Equals("--minimized", StringComparison.OrdinalIgnoreCase));

        _mutex = new System.Threading.Mutex(true, MutexName, out _ownsMutex);
        if (!_ownsMutex)
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

        Task.Run(ListenForSecondInstance);

        MainWin = new MainWindow { StartMinimized = startMinimized };
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
                if (msg == "show") Dispatcher.Invoke(() => MainWin?.BringToFront());
            }
            catch { }
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        if (_ownsMutex)
        {
            try { _mutex?.ReleaseMutex(); } catch { }
        }
        _mutex?.Dispose();
        base.OnExit(e);
    }
}
