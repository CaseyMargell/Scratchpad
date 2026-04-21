using System.Globalization;
using System.Text.RegularExpressions;

namespace Scratchpad;

/// <summary>
/// Evaluates formulas matching the HTML version's behavior exactly.
/// Supports: SUM, AVERAGE/AVG, MIN, MAX, COUNT, COUNTIF, SUMIF, IF, ROUND,
/// ABS, SQRT, POWER, MOD, CONCATENATE, UPPER, LOWER, TRIM, LEN, LEFT, RIGHT,
/// TODAY, NOW, cell references, arithmetic, &amp; concatenation.
/// </summary>
public static class FormulaEvaluator
{
    public static object Evaluate(string raw, GridData grid, HashSet<string> visited)
    {
        if (string.IsNullOrEmpty(raw)) return "";
        var trimmed = raw.Trim();
        if (!trimmed.StartsWith("="))
        {
            if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out var n))
                return n;
            return trimmed;
        }

        var expr = trimmed[1..];

        try
        {
            expr = ReplaceRangeFunctions(expr, grid);
            expr = ReplaceCountifSumif(expr, grid, visited);
            expr = ReplaceIf(expr, grid, visited);
            expr = ReplaceConcatenate(expr, grid, visited);
            expr = ReplaceSingleArg(expr, "ABS", grid, visited,  a => Math.Abs(ToDouble(a)));
            expr = ReplaceSingleArg(expr, "SQRT", grid, visited, a => { var v = ToDouble(a); return v < 0 ? double.NaN : Math.Sqrt(v); });
            expr = ReplaceTwoArg(expr, "ROUND", grid, visited,   (a, b) => { var f = Math.Pow(10, ToDouble(b)); return Math.Round(ToDouble(a) * f) / f; });
            expr = ReplaceTwoArg(expr, "POWER", grid, visited,   (a, b) => Math.Pow(ToDouble(a), ToDouble(b)));
            expr = ReplaceTwoArg(expr, "MOD", grid, visited,     (a, b) => { var x = ToDouble(a); var y = ToDouble(b); return y == 0 ? double.NaN : ((x % y) + y) % y; });
            expr = ReplaceSingleArg(expr, "UPPER", grid, visited, a => JsStr(ToStr(a).ToUpperInvariant()));
            expr = ReplaceSingleArg(expr, "LOWER", grid, visited, a => JsStr(ToStr(a).ToLowerInvariant()));
            expr = ReplaceSingleArg(expr, "TRIM", grid, visited,  a => JsStr(ToStr(a).Trim()));
            expr = ReplaceSingleArg(expr, "LEN", grid, visited,   a => (double)ToStr(a).Length);
            expr = ReplaceTwoArg(expr, "LEFT", grid, visited,    (a, b) => { var s = ToStr(a); var n = (int)ToDouble(b); return JsStr(n >= s.Length ? s : s[..Math.Max(0, n)]); });
            expr = ReplaceTwoArg(expr, "RIGHT", grid, visited,   (a, b) => { var s = ToStr(a); var n = (int)ToDouble(b); return JsStr(n >= s.Length ? s : s[^Math.Max(0, n)..]); });

            expr = Regex.Replace(expr, @"TODAY\(\)", _ => JsStr(DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)), RegexOptions.IgnoreCase);
            expr = Regex.Replace(expr, @"NOW\(\)",   _ => JsStr(DateTime.Now.ToString(CultureInfo.InvariantCulture)), RegexOptions.IgnoreCase);

            // & operator: concatenate text
            expr = Regex.Replace(expr, @"([A-Z]\d+)\s*&\s*", m =>
            {
                var v = ResolveArg(m.Groups[1].Value, grid, new HashSet<string>(visited));
                return JsStr(v?.ToString() ?? "") + "+";
            });

            // Cell references -> numeric values
            expr = Regex.Replace(expr, @"([A-Z])(\d+)", m =>
            {
                var id = m.Value.ToUpperInvariant();
                if (visited.Contains(id)) return "NaN";
                visited.Add(id);
                var raw2 = grid.GetRaw(id);
                var v = string.IsNullOrEmpty(raw2) ? (object)0.0 : Evaluate(raw2, grid, new HashSet<string>(visited));
                return ToDouble(v).ToString("R", CultureInfo.InvariantCulture);
            });

            return EvaluateExpression(expr);
        }
        catch
        {
            return "#ERR";
        }
    }

    // ---- Range-based functions ----

    private static string ReplaceRangeFunctions(string expr, GridData grid)
    {
        expr = Regex.Replace(expr, @"SUM\(([A-Z]\d+):([A-Z]\d+)\)", m =>
        {
            var vals = RangeValues(m.Groups[1].Value, m.Groups[2].Value, grid);
            return vals.OfType<double>().Sum().ToString("R", CultureInfo.InvariantCulture);
        }, RegexOptions.IgnoreCase);

        expr = Regex.Replace(expr, @"(?:AVERAGE|AVG)\(([A-Z]\d+):([A-Z]\d+)\)", m =>
        {
            var nums = RangeValues(m.Groups[1].Value, m.Groups[2].Value, grid).OfType<double>().ToList();
            return (nums.Count == 0 ? 0 : nums.Average()).ToString("R", CultureInfo.InvariantCulture);
        }, RegexOptions.IgnoreCase);

        expr = Regex.Replace(expr, @"(MIN|MAX)\(([A-Z]\d+):([A-Z]\d+)\)", m =>
        {
            var nums = RangeValues(m.Groups[2].Value, m.Groups[3].Value, grid).OfType<double>().ToList();
            if (nums.Count == 0) return "0";
            return (m.Groups[1].Value.Equals("MIN", StringComparison.OrdinalIgnoreCase)
                ? nums.Min() : nums.Max()).ToString("R", CultureInfo.InvariantCulture);
        }, RegexOptions.IgnoreCase);

        expr = Regex.Replace(expr, @"COUNT\(([A-Z]\d+):([A-Z]\d+)\)", m =>
        {
            var n = RangeValues(m.Groups[1].Value, m.Groups[2].Value, grid).OfType<double>().Count();
            return n.ToString(CultureInfo.InvariantCulture);
        }, RegexOptions.IgnoreCase);

        return expr;
    }

    private static string ReplaceCountifSumif(string expr, GridData grid, HashSet<string> visited)
    {
        expr = Regex.Replace(expr, @"COUNTIF\(([A-Z]\d+):([A-Z]\d+)\s*,\s*([^)]+)\)", m =>
        {
            var crit = ResolveArg(m.Groups[3].Value, grid, new HashSet<string>(visited));
            var vals = RangeValues(m.Groups[1].Value, m.Groups[2].Value, grid);
            return vals.Count(v => MatchesCriteria(v, crit)).ToString(CultureInfo.InvariantCulture);
        }, RegexOptions.IgnoreCase);

        expr = Regex.Replace(expr, @"SUMIF\(([A-Z]\d+):([A-Z]\d+)\s*,\s*([^,)]+)(?:\s*,\s*([A-Z]\d+):([A-Z]\d+))?\)", m =>
        {
            var crit = ResolveArg(m.Groups[3].Value, grid, new HashSet<string>(visited));
            var testVals = RangeValues(m.Groups[1].Value, m.Groups[2].Value, grid);
            var sumVals = (m.Groups[4].Success && m.Groups[5].Success)
                ? RangeValues(m.Groups[4].Value, m.Groups[5].Value, grid)
                : testVals;
            double sum = 0;
            for (int i = 0; i < testVals.Count && i < sumVals.Count; i++)
            {
                if (MatchesCriteria(testVals[i], crit) && sumVals[i] is double d) sum += d;
            }
            return sum.ToString("R", CultureInfo.InvariantCulture);
        }, RegexOptions.IgnoreCase);

        return expr;
    }

    private static string ReplaceIf(string expr, GridData grid, HashSet<string> visited)
    {
        return Regex.Replace(expr, @"IF\(([^,]+),([^,]+),([^)]+)\)", m =>
        {
            var cond = m.Groups[1].Value.Trim();
            bool result;
            var opMatch = Regex.Match(cond, @"^(.+?)\s*(>=|<=|<>|!=|>|<|=)\s*(.+)$");
            if (opMatch.Success)
            {
                var left = ResolveArg(opMatch.Groups[1].Value, grid, new HashSet<string>(visited));
                var right = ResolveArg(opMatch.Groups[3].Value, grid, new HashSet<string>(visited));
                var op = opMatch.Groups[2].Value;
                var ln = ToDouble(left); var rn = ToDouble(right);
                bool bothNumeric = !double.IsNaN(ln) && !double.IsNaN(rn);
                result = op switch
                {
                    ">" => bothNumeric && ln > rn,
                    "<" => bothNumeric && ln < rn,
                    ">=" => bothNumeric && ln >= rn,
                    "<=" => bothNumeric && ln <= rn,
                    "=" => bothNumeric ? ln == rn : (ToStr(left) == ToStr(right)),
                    _ => bothNumeric ? ln != rn : (ToStr(left) != ToStr(right))
                };
            }
            else
            {
                var v = ResolveArg(cond, grid, new HashSet<string>(visited));
                result = ToBool(v);
            }
            var branch = result ? m.Groups[2].Value : m.Groups[3].Value;
            var value = ResolveArg(branch, grid, new HashSet<string>(visited));
            return JsonSerialize(value);
        }, RegexOptions.IgnoreCase);
    }

    private static string ReplaceConcatenate(string expr, GridData grid, HashSet<string> visited)
    {
        return Regex.Replace(expr, @"CONCATENATE\(([^)]+)\)", m =>
        {
            var args = SplitArgs(m.Groups[1].Value);
            var parts = args.Select(a => ToStr(ResolveArg(a, grid, new HashSet<string>(visited))));
            return JsStr(string.Concat(parts));
        }, RegexOptions.IgnoreCase);
    }

    private static string ReplaceSingleArg(string expr, string name, GridData grid, HashSet<string> visited, Func<object?, object> fn)
    {
        return Regex.Replace(expr, $@"{name}\(([^)]+)\)", m =>
        {
            var arg = ResolveArg(m.Groups[1].Value, grid, new HashSet<string>(visited));
            var r = fn(arg);
            return FormatForExpression(r);
        }, RegexOptions.IgnoreCase);
    }

    private static string ReplaceTwoArg(string expr, string name, GridData grid, HashSet<string> visited, Func<object?, object?, object> fn)
    {
        return Regex.Replace(expr, $@"{name}\(([^,]+),\s*([^)]+)\)", m =>
        {
            var a = ResolveArg(m.Groups[1].Value, grid, new HashSet<string>(visited));
            var b = ResolveArg(m.Groups[2].Value, grid, new HashSet<string>(visited));
            var r = fn(a, b);
            return FormatForExpression(r);
        }, RegexOptions.IgnoreCase);
    }

    private static string FormatForExpression(object value)
    {
        return value switch
        {
            double d when double.IsNaN(d) => "NaN",
            double d => d.ToString("R", CultureInfo.InvariantCulture),
            string s => JsStr(s),
            bool b => b ? "true" : "false",
            _ => JsStr(value?.ToString() ?? "")
        };
    }

    // ---- Helpers ----

    private static List<object> RangeValues(string start, string end, GridData grid)
    {
        var s = GridData.ParseRef(start.ToUpperInvariant());
        var e = GridData.ParseRef(end.ToUpperInvariant());
        var result = new List<object>();
        if (s == null || e == null) return result;
        int c1 = Math.Min(s.Value.col, e.Value.col), c2 = Math.Max(s.Value.col, e.Value.col);
        int r1 = Math.Min(s.Value.row, e.Value.row), r2 = Math.Max(s.Value.row, e.Value.row);
        for (int c = c1; c <= c2; c++)
            for (int r = r1; r <= r2; r++)
                result.Add(grid.GetValue(GridData.CellId(c, r)) ?? "");
        return result;
    }

    private static object? ResolveArg(string arg, GridData grid, HashSet<string> visited)
    {
        arg = arg.Trim();
        if ((arg.StartsWith("\"") && arg.EndsWith("\"")) ||
            (arg.StartsWith("'") && arg.EndsWith("'")))
            return arg[1..^1];
        var refParsed = GridData.ParseRef(arg.ToUpperInvariant());
        if (refParsed != null)
        {
            var id = GridData.CellId(refParsed.Value.col, refParsed.Value.row);
            if (visited.Contains(id)) return double.NaN;
            visited.Add(id);
            var raw = grid.GetRaw(id);
            return string.IsNullOrEmpty(raw) ? "" : Evaluate(raw, grid, new HashSet<string>(visited));
        }
        if (double.TryParse(arg, NumberStyles.Float, CultureInfo.InvariantCulture, out var n)) return n;
        return arg;
    }

    private static bool MatchesCriteria(object? val, object? crit)
    {
        var c = crit?.ToString() ?? "";
        var m = Regex.Match(c, @"^([><=!]+)(.+)$");
        if (m.Success && double.TryParse(m.Groups[2].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var n))
        {
            if (val is not double dv) return false;
            return m.Groups[1].Value switch
            {
                ">" => dv > n, "<" => dv < n, ">=" => dv >= n, "<=" => dv <= n,
                "<>" or "!=" => dv != n, "=" => dv == n, _ => false
            };
        }
        return string.Equals(val?.ToString(), c, StringComparison.OrdinalIgnoreCase);
    }

    private static List<string> SplitArgs(string s)
    {
        var result = new List<string>();
        int depth = 0;
        var current = new System.Text.StringBuilder();
        foreach (var ch in s)
        {
            if (ch == '(') depth++;
            else if (ch == ')') depth--;
            if (ch == ',' && depth == 0) { result.Add(current.ToString()); current.Clear(); }
            else current.Append(ch);
        }
        result.Add(current.ToString());
        return result;
    }

    private static double ToDouble(object? v)
    {
        if (v is double d) return d;
        if (v is int i) return i;
        if (v is string s && double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var n)) return n;
        return double.NaN;
    }

    private static string ToStr(object? v) => v switch
    {
        null => "",
        double d when d == Math.Floor(d) && !double.IsInfinity(d) => ((long)d).ToString(CultureInfo.InvariantCulture),
        double d => d.ToString("R", CultureInfo.InvariantCulture),
        _ => v.ToString() ?? ""
    };

    private static bool ToBool(object? v)
    {
        if (v is bool b) return b;
        if (v is double d) return d != 0;
        if (v is string s) return !string.IsNullOrEmpty(s);
        return v != null;
    }

    private static string JsStr(string s) => "\"" + s.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";

    private static string JsonSerialize(object? v) => v switch
    {
        null => "\"\"",
        double d => d.ToString("R", CultureInfo.InvariantCulture),
        string s => JsStr(s),
        _ => JsStr(v.ToString() ?? "")
    };

    // ---- Tiny arithmetic evaluator ----
    // Supports: + - * / ( ) unary minus, string literals "..." for concatenation, numbers
    // Returns number for numeric expressions, string for string results, "#ERR" on failure

    private static object EvaluateExpression(string expr)
    {
        var parser = new ExprParser(expr);
        var result = parser.ParseExpression();
        parser.ExpectEnd();
        return result;
    }

    private class ExprParser
    {
        private readonly string _s;
        private int _i;
        public ExprParser(string s) { _s = s; _i = 0; }

        private void Skip() { while (_i < _s.Length && char.IsWhiteSpace(_s[_i])) _i++; }

        public object ParseExpression()
        {
            var left = ParseTerm();
            while (true)
            {
                Skip();
                if (_i >= _s.Length) break;
                var op = _s[_i];
                if (op != '+' && op != '-') break;
                _i++;
                var right = ParseTerm();
                left = ApplyBinary(op, left, right);
            }
            return left;
        }

        private object ParseTerm()
        {
            var left = ParseFactor();
            while (true)
            {
                Skip();
                if (_i >= _s.Length) break;
                var op = _s[_i];
                if (op != '*' && op != '/') break;
                _i++;
                var right = ParseFactor();
                left = ApplyBinary(op, left, right);
            }
            return left;
        }

        private object ParseFactor()
        {
            Skip();
            if (_i >= _s.Length) throw new Exception("Unexpected end");
            var ch = _s[_i];
            if (ch == '-') { _i++; var v = ParseFactor(); return v is double d ? -d : double.NaN; }
            if (ch == '+') { _i++; return ParseFactor(); }
            if (ch == '(')
            {
                _i++;
                var v = ParseExpression();
                Skip();
                if (_i >= _s.Length || _s[_i] != ')') throw new Exception("Missing )");
                _i++;
                return v;
            }
            if (ch == '"')
            {
                _i++;
                var sb = new System.Text.StringBuilder();
                while (_i < _s.Length && _s[_i] != '"')
                {
                    if (_s[_i] == '\\' && _i + 1 < _s.Length) { sb.Append(_s[++_i]); _i++; }
                    else sb.Append(_s[_i++]);
                }
                if (_i < _s.Length) _i++;
                return sb.ToString();
            }
            if (char.IsDigit(ch) || ch == '.')
            {
                int start = _i;
                while (_i < _s.Length && (char.IsDigit(_s[_i]) || _s[_i] == '.' || _s[_i] == 'e' || _s[_i] == 'E' || ((_s[_i] == '+' || _s[_i] == '-') && _i > start && (_s[_i - 1] == 'e' || _s[_i - 1] == 'E')))) _i++;
                var text = _s[start.._i];
                if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var n)) return n;
                return double.NaN;
            }
            if (char.IsLetter(ch))
            {
                int start = _i;
                while (_i < _s.Length && char.IsLetterOrDigit(_s[_i])) _i++;
                var word = _s[start.._i];
                if (word.Equals("NaN", StringComparison.OrdinalIgnoreCase)) return double.NaN;
                if (word.Equals("true", StringComparison.OrdinalIgnoreCase)) return 1.0;
                if (word.Equals("false", StringComparison.OrdinalIgnoreCase)) return 0.0;
                throw new Exception("Unknown token: " + word);
            }
            throw new Exception("Unexpected char: " + ch);
        }

        public void ExpectEnd() { Skip(); if (_i != _s.Length) throw new Exception("Leftover input"); }

        private static object ApplyBinary(char op, object a, object b)
        {
            if (op == '+' && (a is string || b is string))
                return (a?.ToString() ?? "") + (b?.ToString() ?? "");
            double x = a is double da ? da : double.NaN;
            double y = b is double db ? db : double.NaN;
            return op switch
            {
                '+' => x + y,
                '-' => x - y,
                '*' => x * y,
                '/' => y == 0 ? double.NaN : x / y,
                _ => double.NaN
            };
        }
    }
}
