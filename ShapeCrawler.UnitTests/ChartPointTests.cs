using System.Collections.Generic;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Charts;
using ShapeCrawler.UnitTests.Helpers;
using ShapeCrawler.UnitTests.Helpers;
using Xunit;
// ReSharper disable SuggestVarOrType_BuiltInTypes
// ReSharper disable SuggestVarOrType_SimpleTypes

namespace ShapeCrawler.UnitTests;

public class ChartPointTests : ShapeCrawlerTest
{
    [Fact]
    public void Value_Getter_returns_point_value_of_Bar_chart()
    {
        // Arrange
        var pptx21 = GetTestStream("021.pptx");
        var pptx25 = GetTestStream("025_chart.pptx");
        var pres21 = SCPresentation.Open(pptx21);
        var pres25 = SCPresentation.Open(pptx25);
        var shapes1 = pres21.Slides[0].Shapes;
        var chart1 = (IChart) shapes1.First(x => x.Id == 3);
        ISeries chart6Series = ((IChart)pres25.Slides[1].Shapes.First(sp => sp.Id == 4)).SeriesCollection[0];

        // Act
        double pointValue1 = chart1.SeriesCollection[1].Points[0].Value;
        double pointValue2 = chart6Series.Points[0].Value;

        // Assert
        Assert.Equal(56, pointValue1);
        Assert.Equal(72.66, pointValue2);
    }
        
    [Fact]
    public void Value_Getter_returns_point_value_of_Scatter_chart()
    {
        // Arrange
        var pptx = GetTestStream("021.pptx");
        var pres = SCPresentation.Open(pptx);
        var shapes1 = pres.Slides[0].Shapes;
        var chart1 = (IChart) shapes1.First(x => x.Id == 3);
            
        // Act
        double scatterChartPointValue = chart1.SeriesCollection[2].Points[0].Value;
            
        // Assert
        Assert.Equal(44, scatterChartPointValue);
    }
        
    [Fact]
    public void Value_Getter_returns_point_value_of_Line_chart()
    {
        // Arrange
        var chart2 = GetShape<IChart>("021.pptx", 2, 4);
        var point = chart2.SeriesCollection[1].Points[0];

        // Act
        double lineChartPointValue = point.Value;

        // Assert
        Assert.Equal(17.35, lineChartPointValue);
    }
        
    [Fact]
    public void Value_Getter_returns_chart_point()
    {
        // Arrange
        ISeries seriesCase1 = ((IChart)SCPresentation.Open(GetTestStream("021.pptx")).Slides[1].Shapes.First(sp => sp.Id == 3)).SeriesCollection[0];
        ISeries seriesCase2 = ((IChart)SCPresentation.Open(GetTestStream("021.pptx")).Slides[2].Shapes.First(sp => sp.Id == 4)).SeriesCollection[0];
        ISeries seriesCase4 = ((IChart)SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 7)).SeriesCollection[0];

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

    [Theory]
    [MemberData(nameof(TestCasesValueSetter))]
    public void Value_Setter_updates_chart_point(string filename, int slideNumber, string shapeName)
    {
        // Arrange
        var pptxStream = GetTestStream(filename);
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

    [Fact]
    public void Value_Setter_updates_chart_point_in_Embedded_excel_workbook()
    {
        // Arrange
        var pptxStream = GetTestStream("024_chart.pptx");
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

    [Fact]
    public void Value_Getter_returns_chart_point2()
    {
        // Arrange
        var pptxStream = GetTestStream("charts-case004_bars.pptx");
        var pres = SCPresentation.Open(pptxStream);
        {
            var chart = pres.Slides[0].Shapes.First() as IChart;
            var points = chart.SeriesCollection.SelectMany(p => p.Points);
            Assert.Equal(4, chart.SeriesCollection.First().Points.Count());
            Assert.Equal(20, points.Count());
        }
        {
            var chart = pres.Slides[1].Shapes.First() as IChart;
            var points = chart.SeriesCollection.SelectMany(p => p.Points);
            Assert.Equal(4, chart.SeriesCollection.First().Points.Count());
            Assert.Equal(20, points.Count());
        }
        {
            var chart = pres.Slides[2].Shapes.First() as IChart;
            var points = chart.SeriesCollection.SelectMany(p => p.Points);
            Assert.Equal(14, chart.SeriesCollection.Count());
            Assert.Equal(11, chart.SeriesCollection.First().Points.Count());
            Assert.Equal(132, points.Count());
        }
    }

}