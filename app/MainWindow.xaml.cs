using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using Microsoft.Win32;

namespace Scratchpad;

public partial class MainWindow : Window
{
    private HotkeyManager? _hotkey;
    private HwndSource? _hwndSource;
    private bool _quitting;
    private bool _updatingFromGrid;
    public bool StartMinimized { get; set; }

    public MainWindow()
    {
        InitializeComponent();

        // Size and apply theme/grid from settings
        var size = AppSettings.Load().GridSize;
        Grid1.Data.SetDimensions(size.Dimensions().cols, size.Dimensions().rows);

        // Restore window bounds or size to content
        var bounds = AppSettings.Load().WindowBounds;
        if (bounds.HasValue)
        {
            Left = bounds.Value.Left; Top = bounds.Value.Top;
            Width = bounds.Value.Width; Height = bounds.Value.Height;
        }
        else
        {
            SizeToGridContent();
            Left = (SystemParameters.WorkArea.Width - Width) / 2;
            Top = (SystemParameters.WorkArea.Height - Height) / 2;
        }

        ThemeManager.Apply(AppSettings.Load().IsLight);

        Grid1.StatusChanged += msg => FlashStatus(msg);
        Grid1.SelectionChanged += UpdateCellRefAndStats;
        Grid1.FormulaBarTextChanged += t => { if (_updatingFromGrid) return; _updatingFromGrid = true; FormulaBar.Text = t; _updatingFromGrid = false; };

        FormulaBar.PreviewKeyDown += FormulaBar_PreviewKeyDown;
        FormulaBar.TextChanged += FormulaBar_TextChanged;
        FormulaBar.GotKeyboardFocus += (_, _) => { if (!Grid1.IsEditing) Grid1.BeginEditFromFormulaBar(FormulaBar.Text); };

        SourceInitialized += MainWindow_SourceInitialized;
        Closing += MainWindow_Closing;
        StateChanged += (_, _) => { if (WindowState == WindowState.Normal) SaveBoundsSoon(); };
        LocationChanged += (_, _) => SaveBoundsSoon();
        SizeChanged += (_, _) => SaveBoundsSoon();

        Loaded += (_, _) =>
        {
            if (StartMinimized) WindowState = WindowState.Minimized;
            Grid1.SelectCell(AppSettings.Load().LastSelectedCell ?? "A1");
            Grid1.FocusSelectedCell();
            UpdateCellRefAndStats();
        };
    }

    private void SizeToGridContent()
    {
        Width = Grid1.CalculatePreferredContentWidth() + 16;
        Height = Grid1.CalculatePreferredContentHeight() + 110;
    }

    private void MainWindow_SourceInitialized(object? sender, EventArgs e)
    {
        var helper = new WindowInteropHelper(this);
        _hwndSource = HwndSource.FromHwnd(helper.Handle);
        _hwndSource?.AddHook(WndProc);
        _hotkey = new HotkeyManager(helper.Handle);
        var hk = AppSettings.Load().Hotkey;
        _hotkey.Register(hk.Modifiers, hk.Key);
    }

    private const int WM_HOTKEY = 0x0312;
    private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
    {
        if (msg == WM_HOTKEY)
        {
            if (IsVisible && IsActive && WindowState != WindowState.Minimized) Hide();
            else BringToFront();
            handled = true;
        }
        return IntPtr.Zero;
    }

    public void BringToFront()
    {
        Show();
        if (WindowState == WindowState.Minimized) WindowState = WindowState.Normal;
        Activate();
        Topmost = true; Topmost = false;
        Grid1.FocusSelectedCell();
    }

    private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        if (_quitting) { _hotkey?.Dispose(); return; }
        // Hide instead of close, keep hotkey alive
        e.Cancel = true;
        SaveBounds();
        Hide();
    }

    public void Quit()
    {
        _quitting = true;
        SaveBounds();
        _hotkey?.Dispose();
        Close();
        Application.Current.Shutdown();
    }

    private void SaveBoundsSoon()
    {
        // Debounce a bit by just saving on any change; file write is cheap
        SaveBounds();
    }

    private void SaveBounds()
    {
        if (WindowState != WindowState.Normal) return;
        var s = AppSettings.Load();
        s.WindowBounds = new Rect(Left, Top, Width, Height);
        s.LastSelectedCell = Grid1.SelectedCell;
        AppSettings.Save(s);
    }

    // ---- Toolbar ----

    private void SaveBtn_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new SaveFileDialog
        {
            FileName = "scratchpad.xlsx",
            Filter = "Excel Workbook (*.xlsx)|*.xlsx"
        };
        if (dlg.ShowDialog(this) != true) return;
        try
        {
            XlsxService.Save(dlg.FileName, Grid1.Data);
            FlashStatus("Saved");
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Save failed: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OpenBtn_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog { Filter = "Excel Workbook (*.xlsx;*.xls)|*.xlsx;*.xls|CSV (*.csv)|*.csv" };
        if (dlg.ShowDialog(this) != true) return;
        try
        {
            Grid1.ClearAll();
            XlsxService.Load(dlg.FileName, Grid1.Data);
            FlashStatus("Opened: " + Path.GetFileName(dlg.FileName));
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Open failed: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public void ClearAllWithConfirm()
    {
        if (MessageBox.Show(this, "Clear all cells?", "Confirm", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
        Grid1.ClearAll();
    }

    private void SettingsBtn_Click(object sender, RoutedEventArgs e)
    {
        var win = new SettingsWindow(this) { Owner = this };
        win.ShowDialog();
    }

    // ---- Format popover ----

    private void FormatBtn_Click(object sender, RoutedEventArgs e)
    {
        SyncFormatPopup();
        FormatPopup.IsOpen = !FormatPopup.IsOpen;
    }

    private void SyncFormatPopup()
    {
        var f = Grid1.GetSelectionFormat();
        FmtCurrencyBtn.Tag = f?.Style == FormatStyle.Currency ? "active" : null;
        FmtPercentBtn.Tag  = f?.Style == FormatStyle.Percent  ? "active" : null;
        FmtNumberBtn.Tag   = f?.Style == FormatStyle.Number   ? "active" : null;
        FmtDecValue.Text = f?.Decimals?.ToString() ?? "—";
    }

    private void FmtCurrencyBtn_Click(object sender, RoutedEventArgs e) { Grid1.ApplyFormatStyle(FormatStyle.Currency); SyncFormatPopup(); }
    private void FmtPercentBtn_Click(object sender, RoutedEventArgs e)  { Grid1.ApplyFormatStyle(FormatStyle.Percent); SyncFormatPopup(); }
    private void FmtNumberBtn_Click(object sender, RoutedEventArgs e)   { Grid1.ApplyFormatStyle(FormatStyle.Number); SyncFormatPopup(); }
    private void FmtDecLessBtn_Click(object sender, RoutedEventArgs e)  { Grid1.AdjustDecimals(-1); SyncFormatPopup(); }
    private void FmtDecMoreBtn_Click(object sender, RoutedEventArgs e)  { Grid1.AdjustDecimals(+1); SyncFormatPopup(); }
    private void FmtClearBtn_Click(object sender, RoutedEventArgs e)    { Grid1.ClearFormat(); SyncFormatPopup(); }

    private void HelpBtn_Click(object sender, RoutedEventArgs e)
    {
        var win = new HelpWindow { Owner = this };
        win.ShowDialog();
    }

    // ---- Formula bar ----

    private void FormulaBar_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            e.Handled = true;
            Grid1.CommitFormulaBarEdit(FormulaBar.Text);
        }
        else if (e.Key == Key.Escape)
        {
            e.Handled = true;
            Grid1.CancelFormulaBarEdit();
            FormulaBar.Text = Grid1.GetCurrentFormulaText();
            Grid1.FocusSelectedCell();
        }
    }

    private void FormulaBar_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (_updatingFromGrid) return;
        if (!FormulaBar.IsKeyboardFocused) return;
        Grid1.UpdateEditFromFormulaBar(FormulaBar.Text);
    }

    private void UpdateCellRefAndStats()
    {
        CellRefBox.Text = Grid1.SelectedCell;
        _updatingFromGrid = true;
        FormulaBar.Text = Grid1.Data.GetRaw(Grid1.SelectedCell);
        _updatingFromGrid = false;
        var stats = Grid1.GetSelectionStats();
        if (stats.HasValue)
        {
            StatsText.Text = $"Sum: {stats.Value.sum.ToString("N4", CultureInfo.CurrentCulture).TrimEnd('0').TrimEnd('.')}    Avg: {stats.Value.avg.ToString("N4", CultureInfo.CurrentCulture).TrimEnd('0').TrimEnd('.')}    Count: {stats.Value.count}";
        }
        else
        {
            StatsText.Text = "";
        }
    }

    public void ApplyGridSize(GridSize size)
    {
        var (cols, rows) = size.Dimensions();
        var lost = Grid1.Data.CellsOutsideBounds(cols, rows);
        if (lost.Count > 0)
        {
            var resp = MessageBox.Show(this,
                $"Resizing to {size.Label()} will remove {lost.Count} cell{(lost.Count == 1 ? "" : "s")} containing data.\n\nContinue?",
                "Confirm resize", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (resp != MessageBoxResult.Yes) return;
        }
        Grid1.Data.SetDimensions(cols, rows);
        var s = AppSettings.Load();
        s.GridSize = size;
        AppSettings.Save(s);
        SizeToGridContent();
    }

    public bool TryApplyHotkey(HotkeyCombo combo)
    {
        _hotkey?.Unregister();
        var ok = _hotkey?.Register(combo.Modifiers, combo.Key) ?? false;
        if (!ok)
        {
            var old = AppSettings.Load().Hotkey;
            _hotkey?.Register(old.Modifiers, old.Key);
        }
        return ok;
    }

    public void ApplyTheme(bool light)
    {
        ThemeManager.Apply(light);
        var s = AppSettings.Load();
        s.IsLight = light;
        AppSettings.Save(s);
        Grid1.Rebuild();
    }

    private System.Windows.Threading.DispatcherTimer? _statusTimer;
    private void FlashStatus(string msg)
    {
        StatusText.Text = msg;
        _statusTimer?.Stop();
        _statusTimer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromSeconds(1.5) };
        _statusTimer.Tick += (_, _) => { StatusText.Text = "Ready"; _statusTimer!.Stop(); };
        _statusTimer.Start();
    }
}
