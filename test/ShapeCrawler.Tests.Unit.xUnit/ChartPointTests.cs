using System.Collections.Generic;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Charts;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;
// ReSharper disable SuggestVarOrType_BuiltInTypes
// ReSharper disable SuggestVarOrType_SimpleTypes

namespace ShapeCrawler.Tests.Unit;

public class ChartPointTests : SCTest
{
    [Theory]
    [MemberData(nameof(TestCasesValueSetter))]
    public void Value_Setter_updates_chart_point(string filename, int slideNumber, string shapeName)
    {
        // Arrange
        var pptxStream = GetInputStream(filename);
        var pres = SCPresentation.Open(pptxStream);
        var chart = pres.Slides[--slideNumber].Shapes.GetByName<IChart>(shapeName);
        var point = chart.SeriesCollection[0].Points[0];
        const int newChartPointValue = 6;

        // Act
        point.Value = newChartPointValue;

        // Assert
        point.Value.Should().Be(newChartPointValue);

        pres = SaveAndOpenPresentation(pres);
        chart = pres.Slides[slideNumber].Shapes.GetByName<IChart>(shapeName);
        point = chart.SeriesCollection[0].Points[0];
        point.Value.Should().Be(newChartPointValue);
    }

    public static IEnumerable<object[]> TestCasesValueSetter()
    {
        yield return new object[] {"024_chart.pptx", 3, "Chart 4"};
        yield return new object[] {"009_table.pptx", 3, "Chart 5"};
        yield return new object[] {"002.pptx", 1, "Chart 8"};
        yield return new object[] {"021.pptx", 2, "Chart 3"};
        yield return new object[] {"charts-case001.pptx", 1, "chart"};
        yield return new object[] {"charts-case002.pptx", 1, "Chart 1"};
        yield return new object[] {"charts-case003.pptx", 1, "Chart 1"};
    }
}