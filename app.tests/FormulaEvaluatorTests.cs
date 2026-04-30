using Scratchpad;

namespace Scratchpad.Tests;

public class FormulaEvaluatorTests
{
    private static GridData NewGrid() => new();

    private static object Eval(string raw, GridData? g = null)
        => FormulaEvaluator.Evaluate(raw, g ?? NewGrid(), new HashSet<string>());

    // ---- Plain values ----

    [Fact] public void Empty_ReturnsEmpty() => Assert.Equal("", Eval(""));
    [Fact] public void Number_ParsesAsDouble() => Assert.Equal(42.0, Eval("42"));
    [Fact] public void NegativeNumber_ParsesAsDouble() => Assert.Equal(-3.5, Eval("-3.5"));
    [Fact] public void String_ReturnsAsIs() => Assert.Equal("hello", Eval("hello"));

    // ---- Arithmetic ----

    [Theory]
    [InlineData("=1+1", 2.0)]
    [InlineData("=10-3", 7.0)]
    [InlineData("=4*5", 20.0)]
    [InlineData("=10/4", 2.5)]
    [InlineData("=2+3*4", 14.0)]
    [InlineData("=(2+3)*4", 20.0)]
    public void Arithmetic_BasicOperators(string formula, double expected)
        => Assert.Equal(expected, Eval(formula));

    // ---- Cell references ----

    [Fact]
    public void CellRef_ResolvesValue()
    {
        var g = NewGrid();
        g.SetCell("A1", "10");
        Assert.Equal(10.0, Eval("=A1", g));
    }

    [Fact]
    public void CellRef_TransitiveFormula()
    {
        var g = NewGrid();
        g.SetCell("A1", "5");
        g.SetCell("B1", "=A1*2");
        g.SetCell("C1", "=B1+1");
        Assert.Equal(11.0, Eval("=C1", g));
    }

    [Fact]
    public void CycleReference_DoesNotInfiniteLoop()
    {
        var g = NewGrid();
        g.SetCell("A1", "=B1");
        g.SetCell("B1", "=A1");
        // Should produce some sentinel without hanging or throwing
        var v = Eval("=A1", g);
        Assert.NotNull(v);
    }

    // ---- Functions ----

    [Fact]
    public void Sum_AcrossRange()
    {
        var g = NewGrid();
        g.SetCell("A1", "1"); g.SetCell("A2", "2"); g.SetCell("A3", "3");
        Assert.Equal(6.0, Eval("=SUM(A1:A3)", g));
    }

    [Fact]
    public void Sum_IgnoresTextCells()
    {
        var g = NewGrid();
        g.SetCell("A1", "1"); g.SetCell("A2", "hello"); g.SetCell("A3", "3");
        Assert.Equal(4.0, Eval("=SUM(A1:A3)", g));
    }

    [Fact]
    public void Average_OfRange()
    {
        var g = NewGrid();
        g.SetCell("A1", "10"); g.SetCell("A2", "20"); g.SetCell("A3", "30");
        Assert.Equal(20.0, Eval("=AVERAGE(A1:A3)", g));
    }

    [Fact]
    public void If_TrueBranch()
    {
        var g = NewGrid();
        g.SetCell("A1", "10");
        Assert.Equal("big", Eval("=IF(A1>5,\"big\",\"small\")", g));
    }

    [Fact]
    public void If_FalseBranch()
    {
        var g = NewGrid();
        g.SetCell("A1", "1");
        Assert.Equal("small", Eval("=IF(A1>5,\"big\",\"small\")", g));
    }

    [Theory]
    [InlineData("=ROUND(3.14159,2)", 3.14)]
    [InlineData("=ROUND(3.14159,0)", 3.0)]
    [InlineData("=ABS(-7)", 7.0)]
    [InlineData("=SQRT(16)", 4.0)]
    [InlineData("=POWER(2,10)", 1024.0)]
    [InlineData("=MOD(10,3)", 1.0)]
    public void MathFunctions(string formula, double expected)
        => Assert.Equal(expected, Eval(formula));

    [Fact]
    public void Concatenate_JoinsStrings()
    {
        var g = NewGrid();
        g.SetCell("A1", "hello"); g.SetCell("B1", "world");
        Assert.Equal("hello world", Eval("=CONCATENATE(A1,\" \",B1)", g));
    }

    // ============================================================
    //  ExtractReferences — used by the dependency graph
    // ============================================================

    [Fact]
    public void Extract_PlainText_ReturnsEmpty()
    {
        var refs = FormulaEvaluator.ExtractReferences("hello");
        Assert.Empty(refs);
    }

    [Fact]
    public void Extract_NonFormula_ReturnsEmpty()
    {
        // Bare number-looking string with letters, no "="
        var refs = FormulaEvaluator.ExtractReferences("A1");
        Assert.Empty(refs);
    }

    [Fact]
    public void Extract_SimpleRef_ReturnsCell()
    {
        var refs = FormulaEvaluator.ExtractReferences("=A1+5");
        Assert.Contains("A1", refs);
        Assert.Single(refs);
    }

    [Fact]
    public void Extract_MultipleRefs_ReturnsAll()
    {
        var refs = FormulaEvaluator.ExtractReferences("=A1+B2-C3");
        Assert.Contains("A1", refs);
        Assert.Contains("B2", refs);
        Assert.Contains("C3", refs);
        Assert.Equal(3, refs.Count);
    }

    [Fact]
    public void Extract_Range_ExpandsToCells()
    {
        var refs = FormulaEvaluator.ExtractReferences("=SUM(A1:A3)");
        Assert.Contains("A1", refs);
        Assert.Contains("A2", refs);
        Assert.Contains("A3", refs);
        Assert.Equal(3, refs.Count);
    }

    [Fact]
    public void Extract_RectangularRange_ExpandsToAllCells()
    {
        var refs = FormulaEvaluator.ExtractReferences("=SUM(A1:B2)");
        Assert.Contains("A1", refs);
        Assert.Contains("A2", refs);
        Assert.Contains("B1", refs);
        Assert.Contains("B2", refs);
        Assert.Equal(4, refs.Count);
    }

    [Fact]
    public void Extract_MixedRangeAndSingle()
    {
        var refs = FormulaEvaluator.ExtractReferences("=SUM(A1:A3)+B5");
        Assert.Contains("A1", refs);
        Assert.Contains("A2", refs);
        Assert.Contains("A3", refs);
        Assert.Contains("B5", refs);
        Assert.Equal(4, refs.Count);
    }

    // ============================================================
    //  AdjustReferences — used by drag-fill and copy/paste
    //  Regression: copy/paste of =C3-(C4+C5) into column F should
    //  produce =F3-(F4+F5). (Reported by user before v2.1.1.)
    // ============================================================

    [Fact]
    public void Adjust_NonFormula_PassthroughUnchanged()
    {
        var result = FormulaEvaluator.AdjustReferences("hello", 1, 1, 16, 24);
        Assert.Equal("hello", result);
    }

    [Fact]
    public void Adjust_SimpleRef_ShiftsByDelta()
    {
        var result = FormulaEvaluator.AdjustReferences("=A1", 2, 3, 16, 24);
        Assert.Equal("=C4", result);
    }

    [Fact]
    public void Adjust_PreservesOperatorsAndLiterals()
    {
        var result = FormulaEvaluator.AdjustReferences("=A1+5", 1, 0, 16, 24);
        Assert.Equal("=B1+5", result);
    }

    [Fact]
    public void Adjust_Regression_CopyPasteFormula()
    {
        // Bug: copying =C3-(C4+C5) from column C and pasting into column F
        // pasted the evaluated number instead of the adjusted formula. This
        // verifies the adjustment logic that copy/paste relies on.
        var result = FormulaEvaluator.AdjustReferences("=C3-(C4+C5)", 3, 0, 16, 24);
        Assert.Equal("=F3-(F4+F5)", result);
    }

    [Fact]
    public void Adjust_RangeRefs_ShiftsCorrectly()
    {
        var result = FormulaEvaluator.AdjustReferences("=SUM(A1:A3)", 1, 0, 16, 24);
        Assert.Equal("=SUM(B1:B3)", result);
    }

    [Fact]
    public void Adjust_MultipleRanges()
    {
        var result = FormulaEvaluator.AdjustReferences("=SUM(A1:A3)+SUM(B1:B3)", 0, 5, 16, 24);
        Assert.Equal("=SUM(A6:A8)+SUM(B6:B8)", result);
    }

    [Fact]
    public void Adjust_ClampsToGridEdge()
    {
        // Shifting B1 left by 5 would land at column -4; clamp to A.
        var result = FormulaEvaluator.AdjustReferences("=B1", -5, 0, 16, 24);
        Assert.Equal("=A1", result);
    }
}

/// <summary>
/// Tests for the small static helpers exposed by SpreadsheetGrid for which
/// we don't need to instantiate WPF.
/// </summary>
public class SpreadsheetGridHelperTests
{
    /// <summary>
    /// Regression: pressing Shift before pressing Up (i.e. Shift+Up) must not
    /// abort an in-progress arrow-ref selection. The Shift KeyDown fires first
    /// as a separate event, and previously hit the "reset" branch in
    /// EditInput_PreviewKeyDown — so the Up that followed re-initialized
    /// state instead of extending the range.
    /// </summary>
    [Theory]
    [InlineData(System.Windows.Input.Key.LeftShift)]
    [InlineData(System.Windows.Input.Key.RightShift)]
    [InlineData(System.Windows.Input.Key.LeftCtrl)]
    [InlineData(System.Windows.Input.Key.RightCtrl)]
    [InlineData(System.Windows.Input.Key.LeftAlt)]
    [InlineData(System.Windows.Input.Key.RightAlt)]
    [InlineData(System.Windows.Input.Key.LWin)]
    [InlineData(System.Windows.Input.Key.RWin)]
    [InlineData(System.Windows.Input.Key.System)]
    public void ArrowRefReset_ModifierKeys_DoNotReset(System.Windows.Input.Key k)
    {
        Assert.False(InvokeIsArrowRefResetKey(k),
            $"Modifier key {k} must not reset the arrow-ref state.");
    }

    [Theory]
    [InlineData(System.Windows.Input.Key.Up)]
    [InlineData(System.Windows.Input.Key.Down)]
    [InlineData(System.Windows.Input.Key.Left)]
    [InlineData(System.Windows.Input.Key.Right)]
    public void ArrowRefReset_ArrowKeys_DoNotReset(System.Windows.Input.Key k)
    {
        Assert.False(InvokeIsArrowRefResetKey(k),
            $"Arrow key {k} must not reset state — it advances the ref cursor.");
    }

    [Theory]
    [InlineData(System.Windows.Input.Key.A)]
    [InlineData(System.Windows.Input.Key.D5)]
    [InlineData(System.Windows.Input.Key.OemPlus)]
    [InlineData(System.Windows.Input.Key.OemMinus)]
    [InlineData(System.Windows.Input.Key.Space)]
    public void ArrowRefReset_PrintableKeys_DoReset(System.Windows.Input.Key k)
    {
        Assert.True(InvokeIsArrowRefResetKey(k),
            $"Printable key {k} should reset arrow-ref state — user is moving on.");
    }

    private static bool InvokeIsArrowRefResetKey(System.Windows.Input.Key k)
    {
        var method = typeof(SpreadsheetGrid).GetMethod(
            "IsArrowRefResetKey",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        Assert.NotNull(method);
        return (bool)method!.Invoke(null, new object[] { k })!;
    }
}
