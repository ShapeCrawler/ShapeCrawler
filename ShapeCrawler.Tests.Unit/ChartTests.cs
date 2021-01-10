using System.Diagnostics.CodeAnalysis;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Enums;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents.Chart;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.Tests.Unit
{
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    [SuppressMessage("ReSharper", "SuggestVarOrType_BuiltInTypes")]
    public class ChartTests : IClassFixture<PptxFixture>
    {
        private readonly PptxFixture _fixture;

        public ChartTests(PptxFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void XValues_ReturnsParticularXAxisValue_ViaItsCollectionIndexer()
        {
            // Arrange
            IChart chart = _fixture.Pre024.Slides[1].Shapes.First(sp => sp.Id == 5).Chart;

            // Act
            double xValue = chart.XValues[0];

            // Assert
            xValue.Should().Be(10);
            chart.HasXValues.Should().BeTrue();
        }


        [Fact]
        public void Title_ReturnsTextOfTheChartTitle_WhenGetterIsCalled()
        {
            // Arrange
            IChart chart = _fixture.Pre018.Slides[0].Shapes.First(sp => sp.Id == 6).Chart;

            // Act
            string charTitle = chart.Title;

            // Assert
            charTitle.Should().BeEquivalentTo("Test title");
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
            var lineChartPointValue = lineChartSeries.PointValues[0];

            var category = sld1Chart4.Categories[0];

            // Assert
            Assert.Equal(56, barChartPointValue);
            Assert.Equal(44, scatterChartPointValue);
            Assert.Equal(17.35, lineChartPointValue);

            Assert.Equal("2015", category.Name);
        }

        [Fact]
        public void SeriesType_ReturnsChartTypeOfTheSeries()
        {
            // Arrange
            IChart chart = _fixture.Pre021.Slides[0].Shapes.First(sp => sp.Id == 3).Chart;
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
        public void Type_ReturnsChartType()
        {
            // Arrange
            IChart chartCase1 = _fixture.Pre021.Slides[1].Shapes.First(sp => sp.Id == 3).Chart;
            IChart chartCase2 = _fixture.Pre021.Slides[2].Shapes.First(sp => sp.Id == 4).Chart;

            // Act
            ChartType chartTypeCase1 = chartCase1.Type;
            ChartType chartTypeCase2 = chartCase2.Type;

            // Assert
            chartTypeCase1.Should().Be(ChartType.BubbleChart);
            chartTypeCase2.Should().Be(ChartType.ScatterChart);
        }

        [Fact]
        public void SeriesPointValues_ContainsSeriesPointValuesOfTheChart()
        {
            // Arrange
            Series seriesCase1 = _fixture.Pre021.Slides[1].Shapes.First(sp => sp.Id == 3).Chart.SeriesCollection[0];
            Series seriesCase2 = _fixture.Pre021.Slides[2].Shapes.First(sp => sp.Id == 4).Chart.SeriesCollection[0];

            // Act
            double seriesPointValueCase1 = seriesCase1.PointValues[0];
            double seriesPointValueCase2 = seriesCase2.PointValues[0];

            // Arrange
            seriesPointValueCase1.Should().Be(20.4);
            seriesPointValueCase2.Should().Be(2.4);
        }

        [Fact]
        public void Chart4_Test()
        {
            // Arrange
            var pre = new PresentationEx(Properties.Resources._021);
            var sld3Shapes = pre.Slides[2].Shapes;
            var shape4 = sld3Shapes.First(x => x.Id == 4);
            var chart4 = shape4.Chart;

            // Act
            var hasCategories = chart4.HasCategories;

            // Assert
            Assert.False(hasCategories);
        }
    }
}
