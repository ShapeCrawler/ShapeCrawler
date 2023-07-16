using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;
using Assert = NUnit.Framework.Assert;

// ReSharper disable SuggestVarOrType_BuiltInTypes
// ReSharper disable SuggestVarOrType_SimpleTypes

namespace ShapeCrawler.Tests.Unit;

public class ChartPointTests : SCTest
{
    [Test]
    public void Value_Getter_returns_point_value_of_Bar_chart()
    {
        // Arrange
        var pptx21 = GetInputStream("021.pptx");
        var pptx25 = GetInputStream("025_chart.pptx");
        var pres21 = SCPresentation.Open(pptx21);
        var pres25 = SCPresentation.Open(pptx25);
        var shapes1 = pres21.Slides[0].Shapes;
        var chart1 = (IChart)shapes1.First(x => x.Id == 3);
        ISeries chart6Series = ((IChart)pres25.Slides[1].Shapes.First(sp => sp.Id == 4)).SeriesCollection[0];

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
        var pptx = GetInputStream("021.pptx");
        var pres = SCPresentation.Open(pptx);
        var shapes1 = pres.Slides[0].Shapes;
        var chart1 = (IChart)shapes1.First(x => x.Id == 3);

        // Act
        double scatterChartPointValue = chart1.SeriesCollection[2].Points[0].Value;

        // Assert
        Assert.That(scatterChartPointValue, Is.EqualTo(44));
    }

    [Test]
    public void Value_Getter_returns_point_value_of_Line_chart()
    {
        // Arrange
        var chart2 = GetShape<IChart>("021.pptx", 2, 4);
        var point = chart2.SeriesCollection[1].Points[0];

        // Act
        double lineChartPointValue = point.Value;

        // Assert
        Assert.That(lineChartPointValue, Is.EqualTo(17.35));
    }

    [Test]
    public void Value_Getter_returns_chart_point()
    {
        // Arrange
        ISeries seriesCase1 =
            ((IChart)SCPresentation.Open(GetInputStream("021.pptx")).Slides[1].Shapes.First(sp => sp.Id == 3))
            .SeriesCollection[0];
        ISeries seriesCase2 =
            ((IChart)SCPresentation.Open(GetInputStream("021.pptx")).Slides[2].Shapes.First(sp => sp.Id == 4))
            .SeriesCollection[0];
        ISeries seriesCase4 =
            ((IChart)SCPresentation.Open(GetInputStream("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 7))
            .SeriesCollection[0];

        // Act
        double seriesPointValueCase1 = seriesCase1.Points[0].Value;
        double seriesPointValueCase2 = seriesCase2.Points[0].Value;
        double seriesPointValueCase4 = seriesCase4.Points[0].Value;
        double seriesPointValueCase5 = seriesCase4.Points[1].Value;

        // Assert
        seriesPointValueCase1.Should().Be(20.4);
        seriesPointValueCase2.Should().Be(2.4);
        seriesPointValueCase4.Should().Be(8.2);
        seriesPointValueCase5.Should().Be(3.2);
    }

    [Test]
    public void Value_Setter_updates_chart_point_in_Embedded_excel_workbook()
    {
        // Arrange
        var pptxStream = GetInputStream("024_chart.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var chart = pres.Slides[2].Shapes.GetById<IChart>(5);
        var point = chart.SeriesCollection[0].Points[0];
        const int newChartPointValue = 6;

        // Act
        point.Value = newChartPointValue;

        // Assert
        var pointCellValue = GetWorksheetCellValue<double>(chart.WorkbookByteArray, "B2");
        pointCellValue.Should().Be(newChartPointValue);
    }

    [Test]
    public void Value_Getter_returns_chart_point2()
    {
        // Arrange
        var pptxStream = GetInputStream("charts-case004_bars.pptx");
        var pres = SCPresentation.Open(pptxStream);

        var chart1 = pres.Slides[0].Shapes.First() as IChart;
        var points1 = chart1.SeriesCollection.SelectMany(p => p.Points);
        Assert.That(chart1.SeriesCollection.First().Points.Count(), Is.EqualTo(4));
        Assert.That(points1.Count(), Is.EqualTo(20));

        var chart2 = pres.Slides[1].Shapes.First() as IChart;
        var points2 = chart2.SeriesCollection.SelectMany(p => p.Points);
        Assert.That(chart2.SeriesCollection.First().Points.Count(), Is.EqualTo(4));
        Assert.That(points2.Count(), Is.EqualTo(20));

        var chart3 = pres.Slides[2].Shapes.First() as IChart;
        var points3 = chart3.SeriesCollection.SelectMany(p => p.Points);
        Assert.That(chart3.SeriesCollection.Count, Is.EqualTo(14));
        Assert.That(chart3.SeriesCollection.First().Points.Count(), Is.EqualTo(11));
        Assert.That(points3.Count(), Is.EqualTo(132));
    }
}