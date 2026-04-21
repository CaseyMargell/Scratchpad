using System.Globalization;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace Scratchpad;

public partial class SpreadsheetGrid : UserControl
{
    public GridData Data { get; } = new();
    public string SelectedCell { get; private set; } = "A1";
    private (int minC, int maxC, int minR, int maxR)? _selectionRange;
    private string? _selectionEndpoint;
    private const double RowHeaderWidth = 36;
    private const double HeaderHeight = 28;
    private const double DefaultColWidth = 75;
    private const double RowHeight = 28;

    private readonly Dictionary<int, double> _colWidths = new();
    private TextBox? _editingInput;
    private string? _editingCellId;
    private bool _formulaSelecting;
    private int _formulaInsertPos;
    private int _formulaInsertLen;
    private bool _arrowRefActive;
    private int _refCursorCol, _refCursorRow, _refAnchorCol, _refAnchorRow;
    private string? _lastFormulaClickId;

    // Undo/redo
    private readonly Stack<string> _undoStack = new();
    private readonly Stack<string> _redoStack = new();
    private const int MaxUndo = 50;

    public event Action<string>? StatusChanged;
    public event Action? SelectionChanged;

    public SpreadsheetGrid()
    {
        InitializeComponent();
        Data.Changed += Rebuild;
        Loaded += (_, _) => { Rebuild(); FocusSelectedCell(); };
        KeyDown += OnKeyDown;
        PreviewKeyDown += OnPreviewKeyDown;
    }

    public double GetColumnWidth(int col) => _colWidths.TryGetValue(col, out var w) ? w : DefaultColWidth;

    public double CalculatePreferredContentWidth()
    {
        double w = RowHeaderWidth;
        for (int c = 0; c < Data.Cols; c++) w += GetColumnWidth(c);
        return w + 12;
    }

    public double CalculatePreferredContentHeight()
    {
        return HeaderHeight + (Data.Rows * RowHeight) + 12;
    }

    public void Rebuild()
    {
        RootGrid.Children.Clear();
        RootGrid.ColumnDefinitions.Clear();
        RootGrid.RowDefinitions.Clear();

        // Row/Col headers layout
        RootGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(RowHeaderWidth) });
        for (int c = 0; c < Data.Cols; c++)
        {
            RootGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(GetColumnWidth(c)) });
        }

        RootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(HeaderHeight) });
        for (int r = 0; r < Data.Rows; r++)
        {
            RootGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(RowHeight) });
        }

        // Corner
        var corner = new Border
        {
            Background = (Brush)FindResource("BgSurfaceBrush"),
            BorderBrush = (Brush)FindResource("BorderBrush1"),
            BorderThickness = new Thickness(0, 0, 1, 1)
        };
        Grid.SetColumn(corner, 0); Grid.SetRow(corner, 0);
        RootGrid.Children.Add(corner);

        // Column headers
        for (int c = 0; c < Data.Cols; c++)
        {
            var header = new Grid
            {
                Background = (Brush)FindResource("BgSurfaceBrush")
            };
            var border = new Border
            {
                BorderBrush = (Brush)FindResource("BorderBrush1"),
                BorderThickness = new Thickness(0, 0, 1, 1)
            };
            header.Children.Add(border);
            var text = new TextBlock
            {
                Text = ((char)('A' + c)).ToString(),
                Foreground = (Brush)FindResource("TextMutedBrush"),
                FontSize = 13,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            header.Children.Add(text);

            // Column resize grip on the right edge
            var grip = new Border
            {
                Width = 5,
                HorizontalAlignment = HorizontalAlignment.Right,
                Background = Brushes.Transparent,
                Cursor = Cursors.SizeWE,
                Tag = c
            };
            grip.MouseLeftButtonDown += Grip_MouseLeftButtonDown;
            header.Children.Add(grip);

            Grid.SetColumn(header, c + 1); Grid.SetRow(header, 0);
            RootGrid.Children.Add(header);
        }

        // Row headers
        for (int r = 0; r < Data.Rows; r++)
        {
            var rh = new Border
            {
                Background = (Brush)FindResource("BgSurfaceBrush"),
                BorderBrush = (Brush)FindResource("BorderBrush1"),
                BorderThickness = new Thickness(0, 0, 1, 1)
            };
            var text = new TextBlock
            {
                Text = (r + 1).ToString(),
                Foreground = (Brush)FindResource("TextMutedBrush"),
                FontSize = 11,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            rh.Child = text;
            Grid.SetColumn(rh, 0); Grid.SetRow(rh, r + 1);
            RootGrid.Children.Add(rh);
        }

        // Cells
        for (int r = 0; r < Data.Rows; r++)
        {
            for (int c = 0; c < Data.Cols; c++)
            {
                var id = GridData.CellId(c, r);
                var cellBorder = new Border
                {
                    BorderBrush = (Brush)FindResource("BorderBrush1"),
                    BorderThickness = new Thickness(0, 0, 1, 1),
                    Background = (Brush)FindResource("BgBrush"),
                    Tag = id
                };
                var text = new TextBlock
                {
                    Padding = new Thickness(6, 4, 6, 4),
                    Foreground = (Brush)FindResource("TextBrush"),
                    FontSize = 13,
                    VerticalAlignment = VerticalAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis
                };
                var v = Data.GetValue(id);
                text.Text = FormatValue(v);
                if (v is double) text.HorizontalAlignment = HorizontalAlignment.Right;
                cellBorder.Child = text;
                cellBorder.MouseLeftButtonDown += Cell_MouseLeftButtonDown;
                cellBorder.MouseMove += Cell_MouseMove;
                cellBorder.MouseLeftButtonUp += Cell_MouseLeftButtonUp;
                Grid.SetColumn(cellBorder, c + 1); Grid.SetRow(cellBorder, r + 1);
                RootGrid.Children.Add(cellBorder);
            }
        }

        ApplySelectionVisuals();
        SelectionChanged?.Invoke();
    }

    private static string FormatValue(object? v) => v switch
    {
        null => "",
        "" => "",
        double d when double.IsNaN(d) => "#ERR",
        double d when d == Math.Floor(d) && !double.IsInfinity(d) && Math.Abs(d) < 1e15 => ((long)d).ToString("N0", CultureInfo.CurrentCulture),
        double d => d.ToString("0.####", CultureInfo.CurrentCulture),
        _ => v.ToString() ?? ""
    };

    // ---- Selection ----

    public void SelectCell(string id, bool clearRange = true)
    {
        if (GridData.ParseRef(id) == null) return;
        if (clearRange) { _selectionRange = null; _selectionEndpoint = null; }
        SelectedCell = id;
        ApplySelectionVisuals();
        SelectionChanged?.Invoke();
    }

    private void ApplySelectionVisuals()
    {
        foreach (var child in RootGrid.Children)
        {
            if (child is Border b && b.Tag is string id)
            {
                b.BorderThickness = new Thickness(0, 0, 1, 1);
                b.Background = (Brush)FindResource("BgBrush");
            }
        }

        if (_selectionRange is { } rng)
        {
            for (int c = rng.minC; c <= rng.maxC; c++)
                for (int r = rng.minR; r <= rng.maxR; r++)
                {
                    var cellId = GridData.CellId(c, r);
                    var cell = FindCell(cellId);
                    if (cell != null) cell.Background = (Brush)FindResource("RangeBgBrush");
                }
        }

        var sel = FindCell(SelectedCell);
        if (sel != null)
        {
            sel.BorderBrush = (Brush)FindResource("AccentBrush");
            sel.BorderThickness = new Thickness(2);
        }
    }

    private Border? FindCell(string id)
    {
        foreach (var child in RootGrid.Children)
            if (child is Border b && (b.Tag as string) == id) return b;
        return null;
    }

    public void SetSelectionRange(string start, string end)
    {
        var s = GridData.ParseRef(start);
        var e = GridData.ParseRef(end);
        if (s == null || e == null) return;
        _selectionRange = (
            Math.Min(s.Value.col, e.Value.col), Math.Max(s.Value.col, e.Value.col),
            Math.Min(s.Value.row, e.Value.row), Math.Max(s.Value.row, e.Value.row));
        ApplySelectionVisuals();
        SelectionChanged?.Invoke();
    }

    public (int minC, int maxC, int minR, int maxR)? SelectionRange => _selectionRange;

    public void FocusSelectedCell()
    {
        Keyboard.Focus(this);
    }

    // ---- Mouse ----

    private bool _dragging;
    private string? _dragStart;

    private void Cell_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not Border b || b.Tag is not string id) return;
        Focus();
        if (_editingInput != null) { CommitEdit(true); }

        if (IsFormulaMode())
        {
            _formulaSelecting = true;
            var input = _editingInput;
            if (input == null) return;
            var clickedId = id;
            var anchorId = (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) && _lastFormulaClickId != null) ? _lastFormulaClickId : clickedId;
            if (!Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
            {
                _formulaInsertPos = input.SelectionStart;
                _formulaInsertLen = 0;
            }
            var refStr = anchorId == clickedId ? clickedId : anchorId + ":" + clickedId;
            InsertRefIntoFormula(refStr);
            HighlightFormulaRange(anchorId, clickedId);
            if (!Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)) _lastFormulaClickId = clickedId;
            _dragStart = anchorId;
            _dragging = true;
            b.CaptureMouse();
            e.Handled = true;
            return;
        }

        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) && !string.IsNullOrEmpty(SelectedCell))
        {
            SetSelectionRange(SelectedCell, id);
            e.Handled = true;
            return;
        }

        SelectCell(id);
        _dragStart = id;
        _dragging = true;
        b.CaptureMouse();
        e.Handled = true;
    }

    private void Cell_MouseMove(object sender, MouseEventArgs e)
    {
        if (!_dragging || _dragStart == null) return;
        if (e.LeftButton != MouseButtonState.Pressed) return;
        if (sender is not Border b) return;

        // Find the cell under the mouse (within RootGrid)
        var pt = e.GetPosition(RootGrid);
        foreach (var child in RootGrid.Children)
        {
            if (child is Border cellB && cellB.Tag is string overId)
            {
                var topLeft = cellB.TranslatePoint(new Point(0, 0), RootGrid);
                var rect = new Rect(topLeft, new Size(cellB.ActualWidth, cellB.ActualHeight));
                if (rect.Contains(pt))
                {
                    if (IsFormulaMode())
                    {
                        var refStr = overId == _dragStart ? _dragStart : _dragStart + ":" + overId;
                        InsertRefIntoFormula(refStr);
                        HighlightFormulaRange(_dragStart, overId);
                    }
                    else if (overId != _dragStart)
                    {
                        SetSelectionRange(_dragStart, overId);
                    }
                    break;
                }
            }
        }
    }

    private void Cell_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (sender is Border b) b.ReleaseMouseCapture();
        _dragging = false;
        if (_formulaSelecting)
        {
            _formulaSelecting = false;
            _editingInput?.Focus();
            if (_editingInput != null)
            {
                var end = _formulaInsertPos + _formulaInsertLen;
                _editingInput.CaretIndex = end;
            }
        }
    }

    private double _resizeStartX;
    private double _resizeStartW;
    private int _resizeCol = -1;

    private void Grip_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not Border b || b.Tag is not int c) return;
        _resizeCol = c;
        _resizeStartX = e.GetPosition(this).X;
        _resizeStartW = GetColumnWidth(c);
        b.CaptureMouse();
        b.MouseMove += Grip_MouseMove;
        b.MouseLeftButtonUp += Grip_MouseLeftButtonUp;
        e.Handled = true;
    }

    private void Grip_MouseMove(object sender, MouseEventArgs e)
    {
        if (_resizeCol < 0) return;
        var dx = e.GetPosition(this).X - _resizeStartX;
        var newW = Math.Max(40, _resizeStartW + dx);
        _colWidths[_resizeCol] = newW;
        RootGrid.ColumnDefinitions[_resizeCol + 1].Width = new GridLength(newW);
    }

    private void Grip_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (sender is Border b)
        {
            b.ReleaseMouseCapture();
            b.MouseMove -= Grip_MouseMove;
            b.MouseLeftButtonUp -= Grip_MouseLeftButtonUp;
        }
        _resizeCol = -1;
    }

    // ---- Editing ----

    public void StartEdit(string id, string? initialValue = null)
    {
        if (_editingInput != null) CommitEdit(true);
        var cell = FindCell(id);
        if (cell == null) return;
        _editingCellId = id;
        _formulaInsertPos = 0;
        _formulaInsertLen = 0;
        _arrowRefActive = false;

        var tb = new TextBox
        {
            Text = initialValue ?? Data.GetRaw(id),
            Background = (Brush)FindResource("BgBrush"),
            Foreground = (Brush)FindResource("TextBrush"),
            BorderBrush = (Brush)FindResource("AccentBrush"),
            BorderThickness = new Thickness(2),
            FontFamily = new FontFamily("Cascadia Code, Consolas, Courier New"),
            FontSize = 13,
            Padding = new Thickness(4, 2, 4, 2),
            CaretBrush = (Brush)FindResource("TextBrush")
        };
        tb.SelectionStart = tb.Text.Length;
        tb.PreviewKeyDown += EditInput_PreviewKeyDown;
        tb.TextChanged += (_, _) => { _formulaInsertPos = tb.SelectionStart; _formulaInsertLen = 0; FormulaBarTextChanged?.Invoke(tb.Text); };
        tb.LostFocus += (_, _) => { if (!_formulaSelecting && _editingInput == tb) CommitEdit(true); };

        cell.Child = tb;
        _editingInput = tb;
        tb.Focus();
        tb.SelectAll();
        FormulaBarTextChanged?.Invoke(tb.Text);
    }

    public event Action<string>? FormulaBarTextChanged;

    public void CommitEdit(bool save)
    {
        if (_editingInput == null || _editingCellId == null) return;
        var oldRaw = Data.GetRaw(_editingCellId);
        var newText = _editingInput.Text;

        if (save && newText != oldRaw)
        {
            if (newText.TrimStart().StartsWith("=")) newText = AutoCloseParens(newText);
            PushUndo();
            Data.SetCell(_editingCellId, newText);
        }

        // Replace the TextBox with the redrawn cell content
        var cell = FindCell(_editingCellId);
        _editingInput = null;
        _editingCellId = null;
        _arrowRefActive = false;
        ClearFormulaRefs();
        Rebuild();
        ApplySelectionVisuals();
        Focus();
    }

    private static string AutoCloseParens(string text)
    {
        int opens = text.Count(c => c == '(');
        int closes = text.Count(c => c == ')');
        if (opens > closes) return text + new string(')', opens - closes);
        return text;
    }

    public bool IsEditing => _editingInput != null;

    public bool IsFormulaMode()
    {
        return _editingInput != null && _editingInput.Text.TrimStart().StartsWith("=");
    }

    public void UpdateEditFromFormulaBar(string text)
    {
        if (_editingInput == null) return;
        _editingInput.Text = text;
        _editingInput.CaretIndex = text.Length;
    }

    public void BeginEditFromFormulaBar(string text)
    {
        StartEdit(SelectedCell);
        if (_editingInput != null)
        {
            _editingInput.Text = text;
            _editingInput.CaretIndex = text.Length;
        }
    }

    public void CommitFormulaBarEdit(string text)
    {
        PushUndo();
        var raw = text.TrimStart().StartsWith("=") ? AutoCloseParens(text) : text;
        Data.SetCell(SelectedCell, raw);
        Rebuild();
        MoveSelection(0, 1);
        Focus();
    }

    // ---- Formula reference insertion ----

    private void InsertRefIntoFormula(string refStr)
    {
        if (_editingInput == null) return;
        var val = _editingInput.Text;
        var before = val[.._formulaInsertPos];
        var after = val[(_formulaInsertPos + _formulaInsertLen)..];
        _editingInput.Text = before + refStr + after;
        _formulaInsertLen = refStr.Length;
        _editingInput.CaretIndex = _formulaInsertPos + refStr.Length;
        FormulaBarTextChanged?.Invoke(_editingInput.Text);
    }

    private void HighlightFormulaRange(string startId, string endId)
    {
        ClearFormulaRefs();
        var s = GridData.ParseRef(startId);
        var e = GridData.ParseRef(endId);
        if (s == null || e == null) return;
        int c1 = Math.Min(s.Value.col, e.Value.col), c2 = Math.Max(s.Value.col, e.Value.col);
        int r1 = Math.Min(s.Value.row, e.Value.row), r2 = Math.Max(s.Value.row, e.Value.row);
        for (int c = c1; c <= c2; c++)
            for (int r = r1; r <= r2; r++)
            {
                var cell = FindCell(GridData.CellId(c, r));
                if (cell != null)
                {
                    cell.BorderBrush = (Brush)FindResource("FormulaRefBrush");
                    cell.BorderThickness = new Thickness(2);
                }
            }
    }

    private void ClearFormulaRefs()
    {
        foreach (var child in RootGrid.Children)
        {
            if (child is Border b && b.Tag is string)
            {
                b.BorderBrush = (Brush)FindResource("BorderBrush1");
                b.BorderThickness = new Thickness(0, 0, 1, 1);
            }
        }
        ApplySelectionVisuals();
    }

    private bool AtOperatorBoundary()
    {
        if (_editingInput == null) return false;
        var pos = _editingInput.SelectionStart;
        if (pos == 0) return false;
        var before = _editingInput.Text[..pos].TrimEnd();
        if (before.Length == 0) return false;
        return "=+-*/(<>,".Contains(before[^1]);
    }

    // ---- Keyboard ----

    private void EditInput_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            e.Handled = true;
            CommitEdit(true);
            MoveSelection(0, 1);
            Focus();
        }
        else if (e.Key == Key.Tab)
        {
            e.Handled = true;
            CommitEdit(true);
            MoveSelection(Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) ? -1 : 1, 0);
            Focus();
        }
        else if (e.Key == Key.Escape)
        {
            e.Handled = true;
            ClearFormulaRefs();
            _editingInput = null;
            _editingCellId = null;
            Rebuild();
            Focus();
        }
        else if (IsFormulaMode() &&
                 (e.Key == Key.Up || e.Key == Key.Down || e.Key == Key.Left || e.Key == Key.Right))
        {
            if (_arrowRefActive || AtOperatorBoundary())
            {
                e.Handled = true;
                if (!_arrowRefActive)
                {
                    var pos = GridData.ParseRef(SelectedCell);
                    if (pos == null) return;
                    _refCursorCol = pos.Value.col;
                    _refCursorRow = pos.Value.row;
                    _refAnchorCol = _refCursorCol;
                    _refAnchorRow = _refCursorRow;
                    _formulaInsertPos = _editingInput!.SelectionStart;
                    _formulaInsertLen = 0;
                    _arrowRefActive = true;
                }
                if (e.Key == Key.Up) _refCursorRow = Math.Max(0, _refCursorRow - 1);
                else if (e.Key == Key.Down) _refCursorRow = Math.Min(Data.Rows - 1, _refCursorRow + 1);
                else if (e.Key == Key.Left) _refCursorCol = Math.Max(0, _refCursorCol - 1);
                else if (e.Key == Key.Right) _refCursorCol = Math.Min(Data.Cols - 1, _refCursorCol + 1);
                if (!Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
                {
                    _refAnchorCol = _refCursorCol;
                    _refAnchorRow = _refCursorRow;
                }
                var anchorId = GridData.CellId(_refAnchorCol, _refAnchorRow);
                var cursorId = GridData.CellId(_refCursorCol, _refCursorRow);
                var refStr = anchorId == cursorId ? anchorId : anchorId + ":" + cursorId;
                InsertRefIntoFormula(refStr);
                HighlightFormulaRange(anchorId, cursorId);
            }
        }
        else if (_arrowRefActive && e.Key != Key.Up && e.Key != Key.Down && e.Key != Key.Left && e.Key != Key.Right)
        {
            _arrowRefActive = false;
            ClearFormulaRefs();
            var end = _formulaInsertPos + _formulaInsertLen;
            _formulaInsertPos = end;
            _formulaInsertLen = 0;
        }
    }

    private void OnPreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (_editingInput != null) return;

        var ctrl = Keyboard.Modifiers.HasFlag(ModifierKeys.Control);
        var shift = Keyboard.Modifiers.HasFlag(ModifierKeys.Shift);

        // Shift+arrow: extend selection
        if (shift && (e.Key == Key.Up || e.Key == Key.Down || e.Key == Key.Left || e.Key == Key.Right))
        {
            e.Handled = true;
            ExtendSelectionByArrow(e.Key);
            return;
        }

        if (ctrl)
        {
            switch (e.Key)
            {
                case Key.C: e.Handled = true; Copy(false); return;
                case Key.X: e.Handled = true; Copy(true); return;
                case Key.V: e.Handled = true; Paste(); return;
                case Key.A:
                    e.Handled = true;
                    SelectCell("A1", clearRange: false);
                    SetSelectionRange("A1", GridData.CellId(Data.Cols - 1, Data.Rows - 1));
                    return;
                case Key.Z: e.Handled = true; Undo(); return;
                case Key.Y: e.Handled = true; Redo(); return;
            }
        }

        switch (e.Key)
        {
            case Key.Up:    e.Handled = true; MoveSelection(0, -1); return;
            case Key.Down:  e.Handled = true; MoveSelection(0, 1); return;
            case Key.Left:  e.Handled = true; MoveSelection(-1, 0); return;
            case Key.Right: e.Handled = true; MoveSelection(1, 0); return;
            case Key.Tab:   e.Handled = true; MoveSelection(shift ? -1 : 1, 0); return;
            case Key.Enter: e.Handled = true; StartEdit(SelectedCell); return;
            case Key.F2:    e.Handled = true; StartEdit(SelectedCell); return;
            case Key.Delete:
            case Key.Back:
                e.Handled = true;
                DeleteSelection();
                return;
        }
    }

    private void OnKeyDown(object sender, KeyEventArgs e)
    {
        if (_editingInput != null || e.Handled) return;
        if (Keyboard.Modifiers.HasFlag(ModifierKeys.Control) || Keyboard.Modifiers.HasFlag(ModifierKeys.Alt)) return;

        // If a printable character, start edit with that initial value
        var text = TextFromKeyEvent(e);
        if (!string.IsNullOrEmpty(text))
        {
            e.Handled = true;
            StartEdit(SelectedCell, text);
        }
    }

    private static string TextFromKeyEvent(KeyEventArgs e)
    {
        var key = e.Key;
        bool shift = Keyboard.Modifiers.HasFlag(ModifierKeys.Shift);
        // Letters
        if (key >= Key.A && key <= Key.Z) return ((char)('A' + (key - Key.A) + (shift ? 0 : 32))).ToString();
        // Top-row digits (no shift)
        if (!shift && key >= Key.D0 && key <= Key.D9) return ((char)('0' + (key - Key.D0))).ToString();
        if (key >= Key.NumPad0 && key <= Key.NumPad9) return ((char)('0' + (key - Key.NumPad0))).ToString();
        // Common punctuation / symbols
        if (!shift)
        {
            return key switch
            {
                Key.OemMinus => "-",
                Key.OemPlus => "=",
                Key.OemPeriod => ".",
                Key.Decimal => ".",
                Key.OemComma => ",",
                Key.OemQuestion => "/",
                Key.Divide => "/",
                Key.Multiply => "*",
                Key.Add => "+",
                Key.Subtract => "-",
                Key.Space => " ",
                _ => ""
            };
        }
        else
        {
            return key switch
            {
                Key.OemPlus => "+",
                Key.D8 => "*",
                Key.D9 => "(",
                Key.D0 => ")",
                Key.OemQuestion => "?",
                Key.D2 => "@",
                Key.D5 => "%",
                Key.OemComma => "<",
                Key.OemPeriod => ">",
                Key.D6 => "^",
                Key.OemMinus => "_",
                _ => ""
            };
        }
    }

    public void MoveSelection(int dc, int dr)
    {
        var pos = GridData.ParseRef(SelectedCell);
        if (pos == null) { SelectCell("A1"); return; }
        var nc = Math.Clamp(pos.Value.col + dc, 0, Data.Cols - 1);
        var nr = Math.Clamp(pos.Value.row + dr, 0, Data.Rows - 1);
        SelectCell(GridData.CellId(nc, nr));
    }

    private void ExtendSelectionByArrow(Key key)
    {
        var anchor = GridData.ParseRef(SelectedCell);
        if (anchor == null) return;
        var endPt = _selectionEndpoint != null ? GridData.ParseRef(_selectionEndpoint) : anchor;
        if (endPt == null) endPt = anchor;
        int c = endPt.Value.col, r = endPt.Value.row;
        if (key == Key.Up) r = Math.Max(0, r - 1);
        else if (key == Key.Down) r = Math.Min(Data.Rows - 1, r + 1);
        else if (key == Key.Left) c = Math.Max(0, c - 1);
        else if (key == Key.Right) c = Math.Min(Data.Cols - 1, c + 1);
        _selectionEndpoint = GridData.CellId(c, r);
        SetSelectionRange(SelectedCell, _selectionEndpoint);
    }

    private void DeleteSelection()
    {
        PushUndo();
        if (_selectionRange is { } rng)
        {
            for (int c = rng.minC; c <= rng.maxC; c++)
                for (int r = rng.minR; r <= rng.maxR; r++)
                    Data.SetCell(GridData.CellId(c, r), "");
            _selectionRange = null;
        }
        else
        {
            Data.SetCell(SelectedCell, "");
        }
        Rebuild();
    }

    // ---- Clipboard ----

    public void Copy(bool isCut)
    {
        var sb = new StringBuilder();
        var cellsToClear = new List<string>();
        if (_selectionRange is { } rng)
        {
            for (int r = rng.minR; r <= rng.maxR; r++)
            {
                var row = new List<string>();
                for (int c = rng.minC; c <= rng.maxC; c++)
                {
                    var id = GridData.CellId(c, r);
                    var v = Data.GetValue(id);
                    row.Add(FormatValueForClipboard(v));
                    if (isCut && !string.IsNullOrEmpty(Data.GetRaw(id))) cellsToClear.Add(id);
                }
                sb.Append(string.Join("\t", row));
                sb.Append('\n');
            }
        }
        else
        {
            var v = Data.GetValue(SelectedCell);
            sb.Append(FormatValueForClipboard(v));
            if (isCut && !string.IsNullOrEmpty(Data.GetRaw(SelectedCell))) cellsToClear.Add(SelectedCell);
        }
        try { Clipboard.SetText(sb.ToString()); } catch { }
        if (isCut && cellsToClear.Count > 0)
        {
            PushUndo();
            foreach (var id in cellsToClear) Data.SetCell(id, "");
            Rebuild();
        }
        StatusChanged?.Invoke(isCut ? "Cut" : "Copied");
    }

    private static string FormatValueForClipboard(object? v) => v switch
    {
        null => "",
        "" => "",
        double d when d == Math.Floor(d) && !double.IsInfinity(d) => ((long)d).ToString(CultureInfo.InvariantCulture),
        double d => d.ToString("R", CultureInfo.InvariantCulture),
        _ => v.ToString() ?? ""
    };

    public void Paste()
    {
        string text;
        try { text = Clipboard.GetText(); } catch { return; }
        if (string.IsNullOrEmpty(text)) return;

        var pos = GridData.ParseRef(SelectedCell);
        if (pos == null) return;

        PushUndo();
        var lines = text.TrimEnd('\r', '\n').Split('\n');
        for (int r = 0; r < lines.Length && pos.Value.row + r < Data.Rows; r++)
        {
            var line = lines[r].TrimEnd('\r');
            var cols = line.Split('\t');
            for (int c = 0; c < cols.Length && pos.Value.col + c < Data.Cols; c++)
            {
                Data.SetCell(GridData.CellId(pos.Value.col + c, pos.Value.row + r), cols[c]);
            }
        }
        Rebuild();
        StatusChanged?.Invoke("Pasted");
    }

    // ---- Undo/Redo ----

    private void PushUndo()
    {
        var snap = Data.Snapshot();
        if (_undoStack.Count > 0 && _undoStack.Peek() == snap) return;
        _undoStack.Push(snap);
        if (_undoStack.Count > MaxUndo)
        {
            var arr = _undoStack.ToArray();
            _undoStack.Clear();
            for (int i = arr.Length - 2; i >= 0; i--) _undoStack.Push(arr[i]);
        }
        _redoStack.Clear();
    }

    public void Undo()
    {
        if (_undoStack.Count == 0) return;
        _redoStack.Push(Data.Snapshot());
        Data.Restore(_undoStack.Pop());
        Rebuild();
        StatusChanged?.Invoke("Undo");
    }

    public void Redo()
    {
        if (_redoStack.Count == 0) return;
        _undoStack.Push(Data.Snapshot());
        Data.Restore(_redoStack.Pop());
        Rebuild();
        StatusChanged?.Invoke("Redo");
    }

    public void ClearAll()
    {
        PushUndo();
        Data.Clear();
        Rebuild();
        StatusChanged?.Invoke("Cleared");
    }

    // ---- Stats ----

    public (double sum, double avg, int count)? GetSelectionStats()
    {
        var nums = new List<double>();
        if (_selectionRange is { } rng)
        {
            for (int c = rng.minC; c <= rng.maxC; c++)
                for (int r = rng.minR; r <= rng.maxR; r++)
                    if (Data.GetValue(GridData.CellId(c, r)) is double d) nums.Add(d);
        }
        if (nums.Count < 2) return null;
        return (nums.Sum(), nums.Average(), nums.Count);
    }

    public string GetCurrentFormulaText() => Data.GetRaw(SelectedCell);
}
