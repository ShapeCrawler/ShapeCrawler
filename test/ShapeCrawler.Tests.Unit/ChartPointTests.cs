using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Unit.Helpers;
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
        var pptx21 = TestAsset("021.pptx");
        var pptx25 = TestAsset("025_chart.pptx");
        var pres21 = new Presentation(pptx21);
        var pres25 = new Presentation(pptx25);
        var shapes1 = pres21.Slides[0].Shapes;
        var chart1 = (IChart)shapes1.First(x => x.Id == 3);
        ISeries chart6Series = ((IChart)pres25.Slides[1].Shapes.First(sp => sp.Id == 4)).SeriesList[0];

        // Act
        double pointValue1 = chart1.SeriesList[1].Points[0].Value;
        double pointValue2 = chart6Series.Points[0].Value;

        // Assert
        Assert.That(pointValue1, Is.EqualTo(56));
        Assert.That(pointValue2, Is.EqualTo(72.66));
    }

    [Test]
    public void Value_Getter_returns_point_value_of_Scatter_chart()
    {
        // Arrange
        var pptx = TestAsset("021.pptx");
        var pres = new Presentation(pptx);
        var shapes1 = pres.Slides[0].Shapes;
        var chart1 = (IChart)shapes1.First(x => x.Id == 3);

        // Act
        double scatterChartPointValue = chart1.SeriesList[2].Points[0].Value;

        // Assert
        Assert.That(scatterChartPointValue, Is.EqualTo(44));
    }

    [Test]
    public void Value_Getter_returns_point_value_of_Line_chart()
    {
        // Arrange
        var chart2 = GetShape<IChart>("021.pptx", 2, 4);
        var point = chart2.SeriesList[1].Points[0];

        // Act
        double lineChartPointValue = point.Value;

        // Assert
        Assert.That(lineChartPointValue, Is.EqualTo(17.35));
    }

    [Test]
    public void Value_Getter_returns_chart_point()
    {
        // Arrange
        var pres21 = new Presentation(TestAsset("021.pptx"));
        var pres009 = new Presentation(TestAsset("009_table.pptx"));
        var seriesCase1 = pres21.Slides[1].Shapes.GetById<IChart>(3).SeriesList[0];
        var seriesCase2 = pres21.Slides[2].Shapes.GetById<IChart>(4).SeriesList[0];
        var seriesCase3 = pres009.Slides[2].Shapes.GetById<IChart>(7).SeriesList[0];

        // Act
        double seriesPointValueCase1 = seriesCase1.Points[0].Value;
        double seriesPointValueCase2 = seriesCase2.Points[0].Value;
        double seriesPointValueCase4 = seriesCase3.Points[0].Value;
        double seriesPointValueCase5 = seriesCase3.Points[1].Value;

        // Assert
        seriesPointValueCase1.Should().Be(20.4);
        seriesPointValueCase2.Should().Be(2.4);
        seriesPointValueCase4.Should().Be(8.2);
        seriesPointValueCase5.Should().Be(3.2);
    }

    [Test, Ignore("ClosedXML dependency must be removed")]
    public void Value_Setter_updates_chart_point_in_Embedded_excel_workbook()
    {
        // Arrange
        var pres = new Presentation(TestAsset("024_chart.pptx"));
        var chart = pres.Slides[2].Shapes.GetById<IChart>(5);
        var point = chart.SeriesList[0].Points[0];

        // Act
        point.Value = 6;

        // Assert
        var pointCellValue = GetWorksheetCellValue<double>(chart.BookByteArray(), "B2");
        pointCellValue.Should().Be(6);
    }

    [Test]
    public void Value_Getter_returns_chart_point2()
    {
        // Arrange
        var pptxStream = TestAsset("002 bar chart.pptx");
        var pres = new Presentation(pptxStream);

        var chart1 = pres.Slides[0].Shapes.First() as IChart;
        var points1 = chart1.SeriesList.SelectMany(p => p.Points);
        Assert.That(chart1.SeriesList.First().Points.Count(), Is.EqualTo(4));
        Assert.That(points1.Count(), Is.EqualTo(20));

        var chart2 = pres.Slides[1].Shapes.First() as IChart;
        var points2 = chart2.SeriesList.SelectMany(p => p.Points);
        Assert.That(chart2.SeriesList.First().Points.Count(), Is.EqualTo(4));
        Assert.That(points2.Count(), Is.EqualTo(20));

        var chart3 = pres.Slides[2].Shapes.First() as IChart;
        var points3 = chart3.SeriesList.SelectMany(p => p.Points);
        Assert.That(chart3.SeriesList.Count, Is.EqualTo(14));
        Assert.That(chart3.SeriesList.First().Points.Count(), Is.EqualTo(11));
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
        var chart = pres.Slides[--slideNumber].Shapes.GetByName<IChart>(shapeName);
        var point = chart.SeriesList[0].Points[0];
        const int newChartPointValue = 6;

        // Act
        point.Value = newChartPointValue;

        // Assert
        point.Value.Should().Be(newChartPointValue);

        pres = SaveAndOpenPresentation(pres);
        chart = pres.Slides[slideNumber].Shapes.GetByName<IChart>(shapeName);
        point = chart.SeriesList[0].Points[0];
        point.Value.Should().Be(newChartPointValue);
    }
}