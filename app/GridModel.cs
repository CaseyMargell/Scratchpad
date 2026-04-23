namespace Scratchpad;

public enum GridSize { Small, Medium, Large }

public static class GridSizeExtensions
{
    public static (int cols, int rows) Dimensions(this GridSize s) => s switch
    {
        GridSize.Small  => (4, 6),
        GridSize.Medium => (8, 12),
        GridSize.Large  => (16, 24),
        _ => (8, 12)
    };

    public static string Label(this GridSize s) => s switch
    {
        GridSize.Small  => "Small (4 × 6)",
        GridSize.Medium => "Medium (8 × 12)",
        GridSize.Large  => "Large (16 × 24)",
        _ => s.ToString()
    };
}

public class Cell
{
    public string Raw { get; set; } = "";
    public object? Value { get; set; }
}

/// <summary>
/// Describes what changed in a single grid update so subscribers can avoid
/// full rebuilds. <see cref="IsFull"/> means a wholesale change (clear, load,
/// restore, resize) — rebuild everything. Otherwise <see cref="DirtyCells"/>
/// lists exactly which cells need their visuals refreshed.
/// </summary>
public record GridChange(bool IsFull, IReadOnlyCollection<string> DirtyCells);

public class GridData
{
    public int Cols { get; private set; } = 8;
    public int Rows { get; private set; } = 12;
    public Dictionary<string, Cell> Cells { get; } = new();

    // Reverse-dependency graph: dependency -> set of cells whose formulas reference it.
    // When X changes, look up _dependents[X] to find every cell that needs recalc.
    private readonly Dictionary<string, HashSet<string>> _dependents = new();
    // Forward map: cell -> set of cells it currently references. Kept so we can
    // diff and remove stale entries from _dependents when a formula changes.
    private readonly Dictionary<string, HashSet<string>> _dependencies = new();

    public event Action<GridChange>? Changed;

    public void SetDimensions(int cols, int rows)
    {
        Cols = cols;
        Rows = rows;
        var toRemove = Cells.Keys.Where(id =>
        {
            var (c, r) = ParseRef(id) ?? (-1, -1);
            return c < 0 || c >= cols || r < 0 || r >= rows;
        }).ToList();
        foreach (var id in toRemove)
        {
            Cells.Remove(id);
            RemoveFromGraph(id);
        }
        Changed?.Invoke(new GridChange(IsFull: true, Array.Empty<string>()));
    }

    public List<string> CellsOutsideBounds(int newCols, int newRows)
    {
        return Cells.Keys.Where(id =>
        {
            var p = ParseRef(id);
            if (p == null) return false;
            return p.Value.col >= newCols || p.Value.row >= newRows;
        }).ToList();
    }

    public string GetRaw(string id) => Cells.TryGetValue(id, out var c) ? c.Raw : "";
    public object? GetValue(string id) => Cells.TryGetValue(id, out var c) ? c.Value : "";

    /// <summary>
    /// Sets a cell's raw value and recomputes only the dirty subgraph:
    /// this cell plus every cell that transitively depends on it.
    /// </summary>
    public void SetCell(string id, string raw)
    {
        // Update cell storage
        if (string.IsNullOrEmpty(raw))
        {
            Cells.Remove(id);
        }
        else
        {
            if (!Cells.TryGetValue(id, out var cell))
            {
                cell = new Cell();
                Cells[id] = cell;
            }
            cell.Raw = raw;
        }

        // Rebuild dependency graph for this cell only
        UpdateDependencies(id, raw);

        // Dirty set = id + transitive dependents
        var dirty = new HashSet<string>(StringComparer.Ordinal) { id };
        CollectDependents(id, dirty);

        // Recalculate each dirty cell. Order doesn't matter for correctness
        // since Evaluate recurses on demand via the visited-set guard.
        foreach (var d in dirty)
        {
            if (Cells.TryGetValue(d, out var c))
                c.Value = FormulaEvaluator.Evaluate(c.Raw, this, new HashSet<string>());
        }

        Changed?.Invoke(new GridChange(IsFull: false, dirty));
    }

    private void UpdateDependencies(string id, string raw)
    {
        // Drop the cell's old dependencies from the reverse map
        if (_dependencies.TryGetValue(id, out var oldDeps))
        {
            foreach (var dep in oldDeps)
            {
                if (_dependents.TryGetValue(dep, out var set))
                {
                    set.Remove(id);
                    if (set.Count == 0) _dependents.Remove(dep);
                }
            }
            _dependencies.Remove(id);
        }

        // Add the new ones
        var newDeps = FormulaEvaluator.ExtractReferences(raw);
        if (newDeps.Count > 0)
        {
            _dependencies[id] = newDeps;
            foreach (var dep in newDeps)
            {
                if (!_dependents.TryGetValue(dep, out var set))
                {
                    set = new HashSet<string>(StringComparer.Ordinal);
                    _dependents[dep] = set;
                }
                set.Add(id);
            }
        }
    }

    private void RemoveFromGraph(string id)
    {
        UpdateDependencies(id, ""); // clears forward map + reverse entries
        _dependents.Remove(id);       // anyone who depended on this id now dangles; leave entry cleared
    }

    private void CollectDependents(string id, HashSet<string> acc)
    {
        if (!_dependents.TryGetValue(id, out var deps)) return;
        foreach (var d in deps)
        {
            if (acc.Add(d)) CollectDependents(d, acc);
        }
    }

    /// <summary>
    /// Full recalculation — rebuilds the dependency graph and re-evaluates
    /// every cell. Call after bulk changes (Restore, Clear, dimension change).
    /// </summary>
    public void RecalcAll()
    {
        _dependents.Clear();
        _dependencies.Clear();

        foreach (var kv in Cells)
        {
            var deps = FormulaEvaluator.ExtractReferences(kv.Value.Raw);
            if (deps.Count > 0)
            {
                _dependencies[kv.Key] = deps;
                foreach (var dep in deps)
                {
                    if (!_dependents.TryGetValue(dep, out var set))
                    {
                        set = new HashSet<string>(StringComparer.Ordinal);
                        _dependents[dep] = set;
                    }
                    set.Add(kv.Key);
                }
            }
        }

        foreach (var cell in Cells.Values)
            cell.Value = FormulaEvaluator.Evaluate(cell.Raw, this, new HashSet<string>());

        Changed?.Invoke(new GridChange(IsFull: true, Array.Empty<string>()));
    }

    public void Clear()
    {
        Cells.Clear();
        _dependents.Clear();
        _dependencies.Clear();
        Changed?.Invoke(new GridChange(IsFull: true, Array.Empty<string>()));
    }

    public string Snapshot()
    {
        return System.Text.Json.JsonSerializer.Serialize(
            Cells.ToDictionary(kv => kv.Key, kv => kv.Value.Raw));
    }

    public void Restore(string snapshot)
    {
        Cells.Clear();
        _dependents.Clear();
        _dependencies.Clear();
        var map = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(snapshot);
        if (map != null)
        {
            foreach (var kv in map)
                Cells[kv.Key] = new Cell { Raw = kv.Value };
        }
        RecalcAll();
    }

    public static string CellId(int col, int row) => $"{(char)('A' + col)}{row + 1}";

    public static (int col, int row)? ParseRef(string id)
    {
        if (string.IsNullOrEmpty(id) || id.Length < 2) return null;
        if (id[0] < 'A' || id[0] > 'Z') return null;
        if (!int.TryParse(id[1..], out var rowNum) || rowNum < 1) return null;
        return (id[0] - 'A', rowNum - 1);
    }
}
