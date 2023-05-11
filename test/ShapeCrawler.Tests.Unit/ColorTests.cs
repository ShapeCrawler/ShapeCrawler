using FluentAssertions;
using ShapeCrawler.Drawing;
using ShapeCrawler.Tests.Unit.Helpers;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using NUnit.Framework;
using Xunit;

namespace ShapeCrawler.Tests.Unit;

[SuppressMessage("Usage", "xUnit1013:Public method should be marked as test")]
public class ColorTests : SCTest
{
    [Test]
    public void R_G_B_and_Alpha_values_of_White_color()
    {
        var color = SCColor.White;

        // Assert
        color.R.Should().Be(255);
        color.G.Should().Be(255);
        color.B.Should().Be(255);
        color.Alpha.Should().Be(255);
    }

    [Xunit.Theory]
    [MemberData(nameof(HexValuesData))]
    public void FromHex_create_color_from_hexadecimal_code(string hexString, SCColor expectedColor)
    {
        // Act
        var color = SCColor.FromHex(hexString);

        // Assert
        color.R.Should().Be(expectedColor.R);
        color.G.Should().Be(expectedColor.G);
        color.B.Should().Be(expectedColor.B);
        color.Alpha.Should().Be(expectedColor.Alpha);
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
