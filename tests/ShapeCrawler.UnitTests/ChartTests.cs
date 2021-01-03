using System.Diagnostics.CodeAnalysis;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Enums;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents.Chart;
using ShapeCrawler.UnitTests.Helpers;
using Xunit;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.UnitTests
{
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    public class ChartTests : IClassFixture<ReadOnlyTestPresentations>
    {
        private readonly ReadOnlyTestPresentations _fixture;

        public ChartTests(ReadOnlyTestPresentations fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Chart_Test()
        {
            // Arrange
            var pre = _fixture.Pre021;
            var shapes1 = pre.Slides[0].Shapes;
            var shapes2 = pre.Slides[1].Shapes; // TODO: Research why this statement takes mach time
            var sp108 = shapes1.First(x => x.Id == 108);
            var chart3 = shapes1.First(x => x.Id == 3).Chart;
            var sld1Chart4 = shapes1.First(x => x.Id == 4).Chart;
            var sld2Chart4 = shapes2.First(x => x.Id == 4).Chart;
            var lineChartSeries = sld2Chart4.SeriesCollection[1];

            // Act
            var fill = sp108.Fill; //assert: do not throw exception
            var barChartPointValue = chart3.SeriesCollection[1].PointValues[0];
            var scatterChartPointValue = chart3.SeriesCollection[2].PointValues[0];
            var category = sld1Chart4.Categories[0];
            var lineChartPointValue = lineChartSeries.PointValues[0];

            // Assert
            Assert.Equal(56, barChartPointValue);
            Assert.Equal(44, scatterChartPointValue);
            Assert.Equal(17.35, lineChartPointValue);
            Assert.Equal("2015", category.Name);
        }

        [Fact]
        public void SeriesType_ReturnsSeriesChartType()
        {
            // Arrange
            var chart = _fixture.Pre021.Slides[0].Shapes.First(sp => sp.Id == 3).Chart;
            Series series2 = chart.SeriesCollection[1];
            Series series3 = chart.SeriesCollection[2];

            // Act
            ChartType seriesChartType2 = series2.Type;
            ChartType seriesChartType3 = series3.Type;

            // Assert
            seriesChartType2.Should().Be(ChartType.BarChart);
            seriesChartType3.Should().Be(ChartType.ScatterChart);
        }

        [Fact]
        public void Chart3_Test()
        {
            // Arrange
            var pre = new Presentation(Properties.Resources._021);
            var sld2Shapes = pre.Slides[1].Shapes;
            var shape3 = sld2Shapes.First(x => x.Id == 3);
            var chart3 = shape3.Chart;

            // Act
            var pValue = chart3.SeriesCollection[0].PointValues[0];
            var type = chart3.Type;

            // Arrange
            Assert.Equal(20.4, pValue);
            Assert.Equal(ChartType.BubbleChart, type);
        }

        [Fact]
        public void Chart4_Test()
        {
            // Arrange
            var pre = new Presentation(Properties.Resources._021);
            var sld3Shapes = pre.Slides[2].Shapes;
            var shape4 = sld3Shapes.First(x => x.Id == 4);
            var chart4 = shape4.Chart;

            // Act
            var pValue = chart4.SeriesCollection[0].PointValues[0];
            var type = chart4.Type;
            var hasCategories = chart4.HasCategories;

            // Assert
            Assert.Equal(2.4, pValue);
            Assert.Equal(ChartType.ScatterChart, type);
            Assert.False(hasCategories);
        }
    }
}
