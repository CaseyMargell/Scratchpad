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

public class GridData
{
    public int Cols { get; private set; } = 8;
    public int Rows { get; private set; } = 12;
    public Dictionary<string, Cell> Cells { get; } = new();

    public event Action? Changed;

    public void SetDimensions(int cols, int rows)
    {
        Cols = cols;
        Rows = rows;
        // Trim cells outside bounds
        var toRemove = Cells.Keys.Where(id =>
        {
            var (c, r) = ParseRef(id) ?? (-1, -1);
            return c < 0 || c >= cols || r < 0 || r >= rows;
        }).ToList();
        foreach (var id in toRemove) Cells.Remove(id);
        Changed?.Invoke();
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

    public void SetCell(string id, string raw)
    {
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
        RecalcAll();
    }

    public void RecalcAll()
    {
        foreach (var cell in Cells.Values)
            cell.Value = FormulaEvaluator.Evaluate(cell.Raw, this, new HashSet<string>());
        Changed?.Invoke();
    }

    public void Clear()
    {
        Cells.Clear();
        Changed?.Invoke();
    }

    public string Snapshot()
    {
        return System.Text.Json.JsonSerializer.Serialize(
            Cells.ToDictionary(kv => kv.Key, kv => kv.Value.Raw));
    }

    public void Restore(string snapshot)
    {
        Cells.Clear();
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
