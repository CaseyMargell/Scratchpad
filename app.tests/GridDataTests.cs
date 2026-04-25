using Scratchpad;

namespace Scratchpad.Tests;

public class GridDataTests
{
    // ---- Basics ----

    [Fact]
    public void SetCell_StoresRaw()
    {
        var g = new GridData();
        g.SetCell("A1", "hello");
        Assert.Equal("hello", g.GetRaw("A1"));
        Assert.Equal("hello", g.GetValue("A1"));
    }

    [Fact]
    public void SetCell_NumberStoredAsDouble()
    {
        var g = new GridData();
        g.SetCell("A1", "42");
        Assert.Equal(42.0, g.GetValue("A1"));
    }

    [Fact]
    public void SetCell_EmptyValueRemovesCell()
    {
        var g = new GridData();
        g.SetCell("A1", "x");
        g.SetCell("A1", "");
        Assert.Equal("", g.GetRaw("A1"));
        Assert.Empty(g.Cells);
    }

    [Fact]
    public void SetCell_PreservesFormatWhenContentCleared()
    {
        // Regression-prevention: a cell with a format should not lose its
        // format when its value is deleted (so users can pre-format empty cells).
        var g = new GridData();
        g.SetCell("A1", "100");
        g.SetFormat("A1", new CellFormat(FormatStyle.Currency, 2));
        g.SetCell("A1", "");
        Assert.NotNull(g.GetFormat("A1"));
        Assert.Equal(FormatStyle.Currency, g.GetFormat("A1")!.Style);
    }

    // ============================================================
    //  Dependency graph — incremental rebuild correctness
    //  Regression cases: editing a cell must dirty all transitive
    //  dependents, but NOT cells that don't reference it.
    // ============================================================

    [Fact]
    public void DependencyGraph_DirectDependent_Updates()
    {
        var g = new GridData();
        g.SetCell("A1", "5");
        g.SetCell("B1", "=A1*2");
        Assert.Equal(10.0, g.GetValue("B1"));

        g.SetCell("A1", "7");
        Assert.Equal(14.0, g.GetValue("B1"));
    }

    [Fact]
    public void DependencyGraph_TransitiveDependents_Update()
    {
        var g = new GridData();
        g.SetCell("A1", "1");
        g.SetCell("B1", "=A1+1");
        g.SetCell("C1", "=B1+1");
        g.SetCell("D1", "=C1+1");

        g.SetCell("A1", "10");
        Assert.Equal(11.0, g.GetValue("B1"));
        Assert.Equal(12.0, g.GetValue("C1"));
        Assert.Equal(13.0, g.GetValue("D1"));
    }

    [Fact]
    public void DependencyGraph_RangeDependency()
    {
        var g = new GridData();
        g.SetCell("A1", "1"); g.SetCell("A2", "2"); g.SetCell("A3", "3");
        g.SetCell("B1", "=SUM(A1:A3)");
        Assert.Equal(6.0, g.GetValue("B1"));

        g.SetCell("A2", "20");
        Assert.Equal(24.0, g.GetValue("B1"));
    }

    [Fact]
    public void DependencyGraph_FormulaChange_DropsOldDependency()
    {
        // Switch B1 from depending on A1 to depending on A2.
        // A1 changes should no longer dirty B1.
        var g = new GridData();
        g.SetCell("A1", "1");
        g.SetCell("A2", "100");
        g.SetCell("B1", "=A1");
        Assert.Equal(1.0, g.GetValue("B1"));

        g.SetCell("B1", "=A2");
        Assert.Equal(100.0, g.GetValue("B1"));

        // A1 no longer affects B1 — only A2 does.
        g.SetCell("A1", "999");
        Assert.Equal(100.0, g.GetValue("B1"));
        g.SetCell("A2", "200");
        Assert.Equal(200.0, g.GetValue("B1"));
    }

    [Fact]
    public void DependencyGraph_ChangedEvent_DirtyCellsOnIncrement()
    {
        // Regression-prevention: editing a single cell should fire a
        // non-full Changed event with a small dirty set, allowing the UI
        // to update only those cells rather than rebuilding the whole grid.
        var g = new GridData();
        g.SetCell("A1", "1");
        g.SetCell("B1", "=A1+1");

        GridChange? lastChange = null;
        g.Changed += c => lastChange = c;

        g.SetCell("A1", "5");

        Assert.NotNull(lastChange);
        Assert.False(lastChange!.IsFull);
        Assert.Contains("A1", lastChange.DirtyCells);
        Assert.Contains("B1", lastChange.DirtyCells);
    }

    [Fact]
    public void DependencyGraph_UnrelatedCell_NotDirtied()
    {
        var g = new GridData();
        g.SetCell("A1", "1");
        g.SetCell("B1", "=A1");
        g.SetCell("C1", "=99");

        GridChange? lastChange = null;
        g.Changed += c => lastChange = c;
        g.SetCell("A1", "5");

        Assert.NotNull(lastChange);
        Assert.DoesNotContain("C1", lastChange!.DirtyCells);
    }

    // ============================================================
    //  Snapshot / Restore (used by undo/redo and save/load)
    // ============================================================

    [Fact]
    public void Snapshot_Restore_RoundTripsValues()
    {
        var g = new GridData();
        g.SetCell("A1", "1");
        g.SetCell("B1", "=A1+1");
        var snap = g.Snapshot();

        var g2 = new GridData();
        g2.Restore(snap);

        Assert.Equal("1", g2.GetRaw("A1"));
        Assert.Equal("=A1+1", g2.GetRaw("B1"));
        Assert.Equal(2.0, g2.GetValue("B1"));
    }

    [Fact]
    public void Snapshot_Restore_PreservesFormat()
    {
        var g = new GridData();
        g.SetCell("A1", "1234.5");
        g.SetFormat("A1", new CellFormat(FormatStyle.Currency, 2));
        var snap = g.Snapshot();

        var g2 = new GridData();
        g2.Restore(snap);

        var f = g2.GetFormat("A1");
        Assert.NotNull(f);
        Assert.Equal(FormatStyle.Currency, f!.Style);
        Assert.Equal(2, f.Decimals);
    }

    [Fact]
    public void Restore_LegacyFormat_StillWorks()
    {
        // Old undo stacks (pre-format support) used the simpler
        // Dictionary<string,string> snapshot. Restore must tolerate it
        // so existing in-memory undo history doesn't crash after upgrade.
        var legacy = "{\"A1\":\"hello\",\"B1\":\"=A1\"}";

        var g = new GridData();
        g.Restore(legacy);

        Assert.Equal("hello", g.GetRaw("A1"));
        Assert.Equal("=A1", g.GetRaw("B1"));
    }

    // ---- Dimension changes ----

    [Fact]
    public void SetDimensions_RemovesOutOfBoundsCells()
    {
        var g = new GridData();
        g.SetDimensions(8, 12);
        g.SetCell("F8", "x");
        g.SetCell("A1", "y");
        g.SetDimensions(4, 6);
        Assert.Equal("", g.GetRaw("F8"));
        Assert.Equal("y", g.GetRaw("A1"));
    }
}
