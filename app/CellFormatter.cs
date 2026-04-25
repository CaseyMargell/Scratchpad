using System.Globalization;

namespace Scratchpad;

/// <summary>
/// Pure value-to-string formatting for grid cells. Extracted from the UI
/// layer so it can be unit tested without touching WPF.
/// </summary>
public static class CellFormatter
{
    public static string Format(object? value) => Format(value, null, CultureInfo.CurrentCulture);

    public static string Format(object? value, CellFormat? fmt) => Format(value, fmt, CultureInfo.CurrentCulture);

    public static string Format(object? value, CellFormat? fmt, CultureInfo ci)
    {
        if (value is null) return "";
        if (value is string s && s.Length == 0) return "";
        if (value is double dn && double.IsNaN(dn)) return "#ERR";

        // Auto / no explicit format: integer-aware default rendering.
        if (fmt == null || fmt.Style == FormatStyle.Auto)
        {
            return value switch
            {
                double d when d == Math.Floor(d) && !double.IsInfinity(d) && Math.Abs(d) < 1e15
                    => ((long)d).ToString("N0", ci),
                double d => d.ToString("0.####", ci),
                _ => value.ToString() ?? ""
            };
        }

        // Explicit format applies only to numeric values; text stays as-is.
        if (value is not double num) return value.ToString() ?? "";
        if (double.IsInfinity(num)) return "#ERR";

        switch (fmt.Style)
        {
            case FormatStyle.Currency:
            {
                int dec = fmt.Decimals ?? 2;
                var pattern = BuildPattern(dec, fmt.ThousandsSeparator);
                var sign = num < 0 ? "-" : "";
                return sign + "$" + Math.Abs(num).ToString(pattern, ci);
            }
            case FormatStyle.Percent:
            {
                int dec = fmt.Decimals ?? 0;
                var pattern = dec == 0 ? "0" : "0." + new string('0', dec);
                return (num * 100).ToString(pattern, ci) + "%";
            }
            case FormatStyle.Number:
            {
                int dec = fmt.Decimals ?? 2;
                var pattern = BuildPattern(dec, fmt.ThousandsSeparator);
                return num.ToString(pattern, ci);
            }
            default:
                return num.ToString("0.####", ci);
        }
    }

    private static string BuildPattern(int decimals, bool thousands)
    {
        if (decimals == 0) return thousands ? "#,##0" : "0";
        return (thousands ? "#,##0." : "0.") + new string('0', decimals);
    }
}
