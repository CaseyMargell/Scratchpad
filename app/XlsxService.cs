using ClosedXML.Excel;

namespace Scratchpad;

public static class XlsxService
{
    public static void Save(string path, GridData data)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Scratchpad");
        for (int r = 0; r < data.Rows; r++)
        {
            for (int c = 0; c < data.Cols; c++)
            {
                var id = GridData.CellId(c, r);
                var v = data.GetValue(id);
                if (v is null || (v is string s && string.IsNullOrEmpty(s))) continue;
                var xlsxCell = ws.Cell(r + 1, c + 1);
                switch (v)
                {
                    case double d: xlsxCell.Value = d; break;
                    default: xlsxCell.Value = v.ToString(); break;
                }
            }
        }
        wb.SaveAs(path);
    }

    public static void Load(string path, GridData data)
    {
        using var wb = new XLWorkbook(path);
        var ws = wb.Worksheets.First();
        var used = ws.RangeUsed();
        if (used == null) return;
        int lastRow = Math.Min(used.LastRow().RowNumber(), data.Rows);
        int lastCol = Math.Min(used.LastColumn().ColumnNumber(), data.Cols);
        for (int r = 1; r <= lastRow; r++)
        {
            for (int c = 1; c <= lastCol; c++)
            {
                var cell = ws.Cell(r, c);
                var raw = cell.IsEmpty() ? "" : cell.GetString();
                if (!string.IsNullOrEmpty(raw))
                    data.SetCell(GridData.CellId(c - 1, r - 1), raw);
            }
        }
    }
}
