using FluentAssertions;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.CITests;

public class ChartPointTests : SCTest
{
    [Test]
    [TestCase("024_chart.pptx", 3, "Chart 4")]
    [TestCase("009_table.pptx", 3, "Chart 5")]
    [TestCase("002.pptx", 1, "Chart 8")]
    [TestCase("021.pptx", 2, "Chart 3")]
    [TestCase("005 chart.pptx", 1, "chart")]
    [TestCase("004 chart.pptx", 1, "Chart 1")]
    [TestCase("003 chart.pptx", 1, "Chart 1")]
    public void Value_Setter_updates_chart_point(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(file));
        var chart = pres.Slides[--slideNumber].Shape(shapeName).Chart;
        var point = chart.SeriesCollection[0].Points[0];
        const int newChartPointValue = 6;

        // Act
        point.Value = newChartPointValue;

        // Assert
        point.Value.Should().Be(newChartPointValue);

        pres = SaveAndOpenPresentation(pres);
        chart = pres.Slides[slideNumber].Shape(shapeName).Chart;
        point = chart.SeriesCollection[0].Points[0];
        point.Value.Should().Be(newChartPointValue);
    }
    
    [Test]
    public void Value_Getter_returns_point_value_of_Line_chart()
    {
        // Arrange
        var pres = new Presentation(TestAsset("021.pptx"));
        var chart = pres.Slide(2).Shape(4).Chart;
        var point = chart.SeriesCollection[1].Points[0];

        // Act
        double lineChartPointValue = point.Value;

        // Assert
        Assert.That(lineChartPointValue, Is.EqualTo(17.35));
    }
}