using System.Globalization;
using Scratchpad;

namespace Scratchpad.Tests;

public class CellFormatterTests
{
    // Use invariant culture for deterministic output across machines.
    private static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    private static string F(object? v, CellFormat? fmt = null) =>
        CellFormatter.Format(v, fmt, Inv);

    // ---- Default rendering ----

    [Fact] public void Null_ReturnsEmpty() => Assert.Equal("", F(null));
    [Fact] public void EmptyString_ReturnsEmpty() => Assert.Equal("", F(""));
    [Fact] public void NaN_ReturnsErr() => Assert.Equal("#ERR", F(double.NaN));
    [Fact] public void Infinity_ReturnsErr() => Assert.Equal("#ERR", F(double.PositiveInfinity, new CellFormat(FormatStyle.Number)));

    [Fact]
    public void Integer_GetsThousandsSeparatorByDefault()
    {
        Assert.Equal("1,234", F(1234.0));
    }

    [Fact]
    public void Decimal_TrimsTrailingZerosByDefault()
    {
        Assert.Equal("3.14", F(3.14));
    }

    [Fact]
    public void Text_PassesThrough()
    {
        Assert.Equal("hello", F("hello"));
    }

    // ---- Currency ----

    [Theory]
    [InlineData(100.0, 2, "$100.00")]
    [InlineData(1234.567, 2, "$1,234.57")]
    [InlineData(1234.0, 0, "$1,234")]
    [InlineData(0.05, 2, "$0.05")]
    public void Currency_Formats(double v, int dec, string expected)
    {
        Assert.Equal(expected, F(v, new CellFormat(FormatStyle.Currency, dec)));
    }

    [Fact]
    public void Currency_Negative_HasMinusBeforeDollarSign()
    {
        Assert.Equal("-$100.00", F(-100.0, new CellFormat(FormatStyle.Currency, 2)));
    }

    [Fact]
    public void Currency_DefaultDecimalsIsTwo()
    {
        Assert.Equal("$5.00", F(5.0, new CellFormat(FormatStyle.Currency)));
    }

    // ---- Percent ----

    [Theory]
    [InlineData(0.45, 0, "45%")]
    [InlineData(0.4567, 2, "45.67%")]
    [InlineData(1.0, 0, "100%")]
    [InlineData(0.0, 0, "0%")]
    public void Percent_FormatsValueTimes100(double v, int dec, string expected)
    {
        Assert.Equal(expected, F(v, new CellFormat(FormatStyle.Percent, dec)));
    }

    [Fact]
    public void Percent_DefaultDecimalsIsZero()
    {
        Assert.Equal("45%", F(0.45, new CellFormat(FormatStyle.Percent)));
    }

    // ---- Number with thousands ----

    [Fact]
    public void Number_WithThousands()
    {
        Assert.Equal("1,234,567.00", F(1234567.0, new CellFormat(FormatStyle.Number, 2)));
    }

    [Fact]
    public void Number_WithoutThousands()
    {
        Assert.Equal("1234567.00", F(1234567.0, new CellFormat(FormatStyle.Number, 2, false)));
    }

    [Fact]
    public void Number_RoundsToDecimals()
    {
        Assert.Equal("3.14", F(3.14159, new CellFormat(FormatStyle.Number, 2)));
    }

    // ---- Format does not apply to text ----

    [Fact]
    public void Currency_OnTextValue_PassesThrough()
    {
        Assert.Equal("hello", F("hello", new CellFormat(FormatStyle.Currency)));
    }
}
