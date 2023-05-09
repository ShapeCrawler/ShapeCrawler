using FluentAssertions;
using ShapeCrawler.Drawing;
using ShapeCrawler.Tests.Unit.Helpers;
using System.Collections.Generic;
using System.Drawing;
using Xunit;

namespace ShapeCrawler.Tests.Unit;

public class ColorTests : SCTest
{
    [Fact]
    public void Color_IsWhite()
    {
        var color = SCColor.White;

        color.R.Should().Be(255);
        color.G.Should().Be(255);
        color.B.Should().Be(255);
        color.Alpha.Should().Be(255);
    }

    [Theory]
    [MemberData(nameof(HexValuesData))]
    public void Color_Parse_from_hex_values(string hexValue, SCColor expected)
    {
        var rColor = SCColor.TryGetColorFromHex(hexValue, out var color);

        rColor.Should().BeTrue();
        color.R.Should().Be(expected.R);
        color.G.Should().Be(expected.G);
        color.B.Should().Be(expected.B);
        color.Alpha.Should().Be(expected.Alpha);
    }

    [Theory]
    [MemberData(nameof(ColorNamesData))]
    public void Color_Parse_from_names(string name, SCColor expected)
    {
        var color = SCColor.FromName(name);

        color.R.Should().Be(expected.R);
        color.G.Should().Be(expected.G);
        color.B.Should().Be(expected.B);
        color.Alpha.Should().Be(expected.Alpha);
    }

    public static IEnumerable<object[]> HexValuesData => 
        new List<object[]>
        {
            new object[] { "FFF", SCColor.White},
            new object[] { "FFFFFF", SCColor.White},
            new object[] { "FFFFFFFF", SCColor.White},
            new object[] { "0000", SCColor.Transparent},
            new object[] { "00000000", SCColor.Transparent},
            // MS documents colors as ARGB, so we need RGBA and change alpha position.
            // https://learn.microsoft.com/en-us/dotnet/api/system.drawing.color.red?view=net-7.0#system-drawing-color-red
            new object[] { "FF0000FF", new SCColor(Color.Red)},
            // https://learn.microsoft.com/en-us/dotnet/api/system.drawing.color.green?view=net-7.0#system-drawing-color-green
            new object[] { "008000FF", new SCColor(Color.Green)},
        };

    public static IEnumerable<object[]> ColorNamesData =>
        new List<object[]>
        {
            // MS documents colors as ARGB, so we need RGBA and change alpha position.
            // https://learn.microsoft.com/en-us/dotnet/api/system.drawing.color.red?view=net-7.0#system-drawing-color-red
            new object[] { Color.Red.Name, new SCColor(Color.Red)},
            // https://learn.microsoft.com/en-us/dotnet/api/system.drawing.color.green?view=net-7.0#system-drawing-color-green
            new object[] { Color.Silver.Name, new SCColor(Color.Silver)},
        };
}
