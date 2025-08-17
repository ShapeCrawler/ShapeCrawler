using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;
using Assert = NUnit.Framework.Assert;

// ReSharper disable SuggestVarOrType_BuiltInTypes
// ReSharper disable SuggestVarOrType_SimpleTypes

namespace ShapeCrawler.DevTests;

public class ChartPointTests : SCTest
{
    [Test]
    public void Value_Getter_returns_point_value_of_Bar_chart()
    {
        // Arrange
        var pres21 = new Presentation(TestAsset("021.pptx"));
        var pres25 = new Presentation(TestAsset("025_chart.pptx"));
        var shapes1 = pres21.Slides[0].Shapes;
        var chart1 = shapes1.GetById(3).Chart;
        var chart6Series = pres25.Slide(2).Shapes.GetById(4).Chart.SeriesCollection[0];

        // Act
        double pointValue1 = chart1.SeriesCollection[1].Points[0].Value;
        double pointValue2 = chart6Series.Points[0].Value;

        // Assert
        Assert.That(pointValue1, Is.EqualTo(56));
        Assert.That(pointValue2, Is.EqualTo(72.66));
    }

    [Test]
    public void Value_Getter_returns_point_value_of_Scatter_chart()
    {
        // Arrange
        var pres = new Presentation(TestAsset("021.pptx"));
        var chart = pres.Slide(1).Shape(3).Chart;

        // Act
        double scatterChartPointValue = chart.SeriesCollection[2].Points[0].Value;

        // Assert
        Assert.That(scatterChartPointValue, Is.EqualTo(44));
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

    [Test]
    public void Value_Getter_returns_chart_point()
    {
        // Arrange
        var pres1 = new Presentation(TestAsset("021.pptx"));
        var pres2 = new Presentation(TestAsset("009_table.pptx"));
        var series1 = pres1.Slide(2).Shape(3).Chart.SeriesCollection[0];
        var series2 = pres1.Slide(3).Shape(4).Chart.SeriesCollection[0];
        var series3 = pres2.Slide(3).Shape(7).Chart.SeriesCollection[0];

        // Act
        double seriesPointValue1 = series1.Points[0].Value;
        double seriesPointValue2 = series2.Points[0].Value;
        double seriesPointValue4 = series3.Points[0].Value;
        double seriesPointValue5 = series3.Points[1].Value;

        // Assert
        seriesPointValue1.Should().Be(20.4);
        seriesPointValue2.Should().Be(2.4);
        seriesPointValue4.Should().Be(8.2);
        seriesPointValue5.Should().Be(3.2);
    }

    [Test, Ignore("ClosedXML dependency must be removed")]
    public void Value_Setter_updates_chart_point_in_Embedded_excel_workbook()
    {
        // Arrange
        var pres = new Presentation(TestAsset("024_chart.pptx"));
        var chart = pres.Slides[2].Shape(5).Chart;
        var point = chart.SeriesCollection[0].Points[0];

        // Act
        point.Value = 6;

        // Assert
        var pointCellValue = GetWorksheetCellValue<double>(chart.GetWorksheetByteArray(), "B2");
        pointCellValue.Should().Be(6);
    }

    [Test]
    public void Value_Getter_returns_chart_point2()
    {
        // Arrange
        var pres = new Presentation(TestAsset("002 bar chart.pptx"));

        var chart1 = pres.Slides[0].Shapes.First().Chart;
        var points1 = chart1.SeriesCollection.SelectMany(p => p.Points);
        Assert.That(chart1.SeriesCollection.First().Points.Count, Is.EqualTo(4));
        Assert.That(points1.Count(), Is.EqualTo(20));

        var chart2 = pres.Slides[1].Shapes.First().Chart;
        var points2 = chart2.SeriesCollection.SelectMany(p => p.Points);
        Assert.That(chart2.SeriesCollection.First().Points.Count(), Is.EqualTo(4));
        Assert.That(points2.Count(), Is.EqualTo(20));

        var chart3 = pres.Slides[2].Shapes.First().Chart;
        var points3 = chart3.SeriesCollection.SelectMany(p => p.Points);
        Assert.That(chart3.SeriesCollection.Count, Is.EqualTo(14));
        Assert.That(chart3.SeriesCollection.First().Points.Count(), Is.EqualTo(11));
        Assert.That(points3.Count(), Is.EqualTo(132));
    }
    
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
}