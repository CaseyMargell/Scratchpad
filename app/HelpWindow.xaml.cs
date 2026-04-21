using System.Diagnostics;
using System.Windows;

namespace Scratchpad;

public partial class HelpWindow : Window
{
    public HelpWindow() { InitializeComponent(); }

    private void GithubBtn_Click(object sender, RoutedEventArgs e)
        => OpenUrl("https://github.com/CaseyMargell/Scratchpad");

    private void LinkedInBtn_Click(object sender, RoutedEventArgs e)
        => OpenUrl("https://www.linkedin.com/in/cmargell");

    private static void OpenUrl(string url)
    {
        try { Process.Start(new ProcessStartInfo { FileName = url, UseShellExecute = true }); }
        catch { }
    }
}
