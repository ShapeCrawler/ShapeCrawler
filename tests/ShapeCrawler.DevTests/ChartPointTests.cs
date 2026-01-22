using System;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;
using Assert = NUnit.Framework.Assert;

// ReSharper disable SuggestVarOrType_BuiltInTypes
// ReSharper disable SuggestVarOrType_SimpleTypes

namespace ShapeCrawler.DevTests;

public class ChartPointTests : SCTest
{
    private static IChart GetChart(IShape shape)
    {
        return (IChart?)shape.BarChart
               ?? (IChart?)shape.ColumnChart
               ?? (IChart?)shape.LineChart
               ?? (IChart?)shape.PieChart
               ?? (IChart?)shape.ScatterChart
               ?? (IChart?)shape.BubbleChart
               ?? (IChart?)shape.AreaChart
               ?? throw new InvalidOperationException("The shape doesn't contain a chart.");
    }

    [Test]
    public void Value_Getter_returns_point_value()
    {
        // Arrange
        var pres1 = new Presentation(TestAsset("021.pptx"));
        var pres2 = new Presentation(TestAsset("025_chart.pptx"));
        var shapes1 = pres1.Slides[0].Shapes;
        var scatterChart = shapes1.GetById(3).ScatterChart;
        var chart6Series = pres2.Slide(2).Shapes.GetById(4).BarChart.SeriesCollection[0];

        // Act
        double pointValue1 = scatterChart.SeriesCollection[1].Points[0].Value;
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
        var chart = pres.Slide(1).Shape(3).ScatterChart;

        // Act
        double scatterChartPointValue = chart.SeriesCollection[2].Points[0].Value;

        // Assert
        Assert.That(scatterChartPointValue, Is.EqualTo(44));
    }
    
    [Test]
    public void Value_Getter_returns_chart_point()
    {
        // Arrange
        var pres1 = new Presentation(TestAsset("021.pptx"));
        var pres2 = new Presentation(TestAsset("009_table.pptx"));
        var series1 = GetChart(pres1.Slide(2).Shape(3)).SeriesCollection[0];
        var series2 = GetChart(pres1.Slide(3).Shape(4)).SeriesCollection[0];
        var series3 = GetChart(pres2.Slide(3).Shape(7)).SeriesCollection[0];

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
        var chart = GetChart(pres.Slides[2].Shape(5));
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

        var chart1 = GetChart(pres.Slides[0].Shapes.First());
        var points1 = chart1.SeriesCollection.SelectMany(p => p.Points);
        Assert.That(chart1.SeriesCollection.First().Points.Count, Is.EqualTo(4));
        Assert.That(points1.Count(), Is.EqualTo(20));

        var chart2 = GetChart(pres.Slides[1].Shapes.First());
        var points2 = chart2.SeriesCollection.SelectMany(p => p.Points);
        Assert.That(chart2.SeriesCollection.First().Points.Count(), Is.EqualTo(4));
        Assert.That(points2.Count(), Is.EqualTo(20));

        var chart3 = GetChart(pres.Slides[2].Shapes.First());
        var points3 = chart3.SeriesCollection.SelectMany(p => p.Points);
        Assert.That(chart3.SeriesCollection.Count, Is.EqualTo(14));
        Assert.That(chart3.SeriesCollection.First().Points.Count(), Is.EqualTo(11));
        Assert.That(points3.Count(), Is.EqualTo(132));
    }
}
