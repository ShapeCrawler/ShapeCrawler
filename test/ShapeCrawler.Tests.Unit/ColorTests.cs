using FluentAssertions;
using ShapeCrawler.Drawing;
using ShapeCrawler.Tests.Unit.Helpers;
using System.Collections.Generic;
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

    public static IEnumerable<object[]> HexValuesData => 
        new List<object[]>
        {
            new object[] { "FFF", SCColor.White},
            new object[] { "FFFFFF", SCColor.White},
            new object[] { "FFFFFFFF", SCColor.White},
            new object[] { "0000", SCColor.Transparent},
            new object[] { "00000000", SCColor.Transparent},
        };
}
