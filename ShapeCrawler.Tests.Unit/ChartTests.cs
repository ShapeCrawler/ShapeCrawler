using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Charts;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Properties;
using Xunit;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.Tests.Unit
{
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    [SuppressMessage("ReSharper", "SuggestVarOrType_BuiltInTypes")]
    public class ChartTests : IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public ChartTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void XValues_ReturnsParticularXAxisValue_ViaItsCollectionIndexer()
        {
            // Arrange
            IChart chart = _fixture.Pre024.Slides[1].Shapes.First(sp => sp.Id == 5) as IChart;

            // Act
            double xValue = chart.XValues[0];

            // Assert
            xValue.Should().Be(10);
            chart.HasXValues.Should().BeTrue();
        }


        [Fact]
        public void HasXValues()
        {
            // Arrange
            SCSlide slide1 = _fixture.Pre025.Slides[0];
            SCSlide slide2 = _fixture.Pre025.Slides[1];
            IChart chart8 = slide1.Shapes.First(x => x.Id == 8) as IChart;
            IChart chart11 = slide2.Shapes.First(x => x.Id == 11) as IChart;

            // Act
            var chart8HasXValues = chart8.HasXValues;
            var chart11HasXValues = chart11.HasXValues;

            // Assert
            Assert.False(chart8HasXValues);
            Assert.False(chart11HasXValues);
        }

        [Fact]
        public void HasCategories_ReturnsFalse_WhenAChartHasNotCategories()
        {
            // Arrange
            IChart chart = (IChart)_fixture.Pre021.Slides[2].Shapes.First(sp => sp.Id == 4);

            // Act
            bool hasChartCategories = chart.HasCategories;

            // Assert
            hasChartCategories.Should().BeFalse();
        }

        [Fact]
        public void TitleAndHasTitle_ReturnChartTitleStringAndFlagIndicatingWhetherChartHasATitle()
        {
            // Arrange
            IChart chartCase1 = (IChart)_fixture.Pre018.Slides[0].Shapes.First(sp => sp.Id == 6);
            IChart chartCase2 = (IChart)_fixture.Pre025.Slides[0].Shapes.First(sp => sp.Id == 7);
            IChart chartCase3 = (IChart)_fixture.Pre013.Slides[0].Shapes.First(sp => sp.Id == 5);
            IChart chartCase4 = (IChart)_fixture.Pre013.Slides[0].Shapes.First(sp => sp.Id == 4);
            IChart chartCase5 = (IChart)_fixture.Pre019.Slides[0].Shapes.First(sp => sp.Id == 4);
            IChart chartCase6 = (IChart)_fixture.Pre013.Slides[0].Shapes.First(sp => sp.Id == 6);
            IChart chartCase7 = (IChart)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 7);
            IChart chartCase8 = (IChart)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 6);
            IChart chartCase9 = (IChart)_fixture.Pre009.Slides[4].Shapes.First(sp => sp.Id == 6);
            IChart chartCase10 = (IChart)_fixture.Pre009.Slides[4].Shapes.First(sp => sp.Id == 3);
            IChart chartCase11 = (IChart)_fixture.Pre009.Slides[4].Shapes.First(sp => sp.Id == 5);
            
            // Act
            string charTitleCase1 = chartCase1.Title;
            string charTitleCase2 = chartCase2.Title;
            string charTitleCase3 = chartCase3.Title;
            string charTitleCase5 = chartCase5.Title;
            string charTitleCase7 = chartCase7.Title;
            string charTitleCase8 = chartCase8.Title;
            string charTitleCase9 = chartCase9.Title;
            string charTitleCase10 = chartCase10.Title;
            string charTitleCase11 = chartCase11.Title;
            bool hasTitleCase4 = chartCase4.HasTitle;
            bool hasTitleCase6 = chartCase6.HasTitle;

            // Assert
            charTitleCase1.Should().BeEquivalentTo("Test title");
            charTitleCase2.Should().BeEquivalentTo("Series 1_id7");
            charTitleCase3.Should().BeEquivalentTo("Title text");
            charTitleCase5.Should().BeEquivalentTo("Test title");
            charTitleCase7.Should().BeEquivalentTo("Sales");
            charTitleCase8.Should().BeEquivalentTo("Sales2");
            charTitleCase9.Should().BeEquivalentTo("Sales3");
            charTitleCase10.Should().BeEquivalentTo("Sales4");
            charTitleCase11.Should().BeEquivalentTo("Sales5");
            hasTitleCase4.Should().BeFalse();
            hasTitleCase6.Should().BeFalse();
        }

        [Fact]
        public void Chart_Test()
        {
            // Arrange
            IPresentation presentation = _fixture.Pre021;
            var shapes1 = presentation.Slides[0].Shapes;
            var shapes2 = presentation.Slides[1].Shapes; 
            var chart3 = shapes1.First(x => x.Id == 3) as IChart;
            var sld2Chart4 = shapes2.First(x => x.Id == 4) as IChart;
            var lineChartSeries = sld2Chart4.SeriesCollection[1];

            // Act
            var barChartPointValue = chart3.SeriesCollection[1].PointValues[0];
            var scatterChartPointValue = chart3.SeriesCollection[2].PointValues[0];
            IReadOnlyList<double> pointValues = lineChartSeries.PointValues;
            var lineChartPointValue = pointValues[0];

            // Assert
            Assert.Equal(56, barChartPointValue);
            Assert.Equal(44, scatterChartPointValue);
            Assert.Equal(17.35, lineChartPointValue);
        }

        [Fact]
        public void CategoryName_GetterReturnsChartCategoryName()
        {
            // Arrange
            IChart chartCase1 = (IChart)_fixture.Pre025.Slides[0].Shapes.First(sp => sp.Id == 4);
            IChart chartCase2 = (IChart)_fixture.Pre021.Slides[0].Shapes.First(sp => sp.Id == 4);
            IChart chartCase3 = (IChart)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 7);

            // Act-Assert
            chartCase1.Categories[0].Name.Should().BeEquivalentTo("Dresses");
            chartCase2.Categories[0].Name.Should().BeEquivalentTo("2015");
            chartCase3.Categories[0].Name.Should().BeEquivalentTo("Q1");
            chartCase3.Categories[1].Name.Should().BeEquivalentTo("Q2");
            chartCase3.Categories[2].Name.Should().BeEquivalentTo("Q3");
            chartCase3.Categories[3].Name.Should().BeEquivalentTo("Q4");
        }

        [Fact]
        public void CategoryName_GetterReturnsChartCategoryName_OfMultiCategoryChart()
        {
            // Arrange
            IChart chartCase1 = (IChart)_fixture.Pre025.Slides[0].Shapes.First(sp => sp.Id == 4);

            // Act-Assert
            chartCase1.Categories[0].MainCategory.Name.Should().BeEquivalentTo("Clothing");
        }

#if DEBUG
        [Fact]
        public void CategoryName_SetterChangeName_OfCategoryInPieChart()
        {
            // Arrange
            IPresentation presentation = SCPresentation.Open(Resources._025, true);
            MemoryStream mStream = new();
            IChart pieChart4 = (IChart)presentation.Slides[0].Shapes.First(sp => sp.Id == 7);
            const string newCategoryName = "Category 1_new";

            // Act
            pieChart4.Categories[0].Name = newCategoryName;

            // Assert
            pieChart4.Categories[0].Name.Should().Be(newCategoryName);
            presentation.SaveAs(mStream);
            presentation = SCPresentation.Open(mStream, false);
            pieChart4 = (IChart)presentation.Slides[0].Shapes.First(sp => sp.Id == 7);
            pieChart4.Categories[0].Name.Should().Be(newCategoryName);
        }

        [Fact(Skip = "In Progress")]
        public void CategoryName_SetterChangeName_OfMainCategoryInMultiLevelCategoryBarChart()
        {
            // Arrange
            IPresentation presentation = SCPresentation.Open(Resources._025, true);
            MemoryStream mStream = new();
            IChart barChart2 = (IChart)_fixture.Pre025.Slides[0].Shapes.First(sp => sp.Id == 4);
            const string newMainCategoryName = "Clothing_new";

            // Act
            barChart2.Categories[0].MainCategory.Name = newMainCategoryName;

            // Assert
            barChart2.Categories[0].Name.Should().Be(newMainCategoryName);
            presentation.SaveAs(mStream);
            presentation = SCPresentation.Open(mStream, false);
            barChart2 = (IChart)presentation.Slides[0].Shapes.First(sp => sp.Id == 4);
            barChart2.Categories[0].Name.Should().Be(newMainCategoryName);
        }
#endif
        [Fact]
        public void SeriesType_ReturnsChartTypeOfTheSeries()
        {
            // Arrange
            IChart chart = (IChart)_fixture.Pre021.Slides[0].Shapes.First(sp => sp.Id == 3);
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
        public void SeriesCollection_CounterReturnsNumberOfTheSeriesOnTheChart()
        {
            // Arrange
            IChart chartCase1 = (IChart)_fixture.Pre013.Slides[0].Shapes.First(sp => sp.Id == 5);
            IChart chartCase2 = (IChart)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 7);
            
            // Act
            int seriesCountCase1 = chartCase1.SeriesCollection.Count;
            int seriesCountCase2 = chartCase2.SeriesCollection.Count;

            // Assert
            seriesCountCase1.Should().Be(3);
            seriesCountCase2.Should().Be(1);
        }

        [Fact]
        public void SeriesPointValue_ReturnsChartSeriesPointValue()
        {
            // Arrange
            Series seriesCase1 = ((IChart)_fixture.Pre021.Slides[1].Shapes.First(sp => sp.Id == 3)).SeriesCollection[0];
            Series seriesCase2 = ((IChart)_fixture.Pre021.Slides[2].Shapes.First(sp => sp.Id == 4)).SeriesCollection[0];
            Series seriesCase3 = ((IChart)_fixture.Pre025.Slides[1].Shapes.First(sp => sp.Id == 4)).SeriesCollection[0];
            Series seriesCase4 = ((IChart)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 7)).SeriesCollection[0];

            // Act
            double seriesPointValueCase1 = seriesCase1.PointValues[0];
            double seriesPointValueCase2 = seriesCase2.PointValues[0];
            double seriesPointValueCase3 = seriesCase3.PointValues[0];
            double seriesPointValueCase4 = seriesCase4.PointValues[0];
            double seriesPointValueCase5 = seriesCase4.PointValues[1];

            // Assert
            seriesPointValueCase1.Should().Be(20.4);
            seriesPointValueCase2.Should().Be(2.4);
            seriesPointValueCase3.Should().Be(72.7);
            seriesPointValueCase4.Should().Be(8.2);
            seriesPointValueCase5.Should().Be(3.2);
        }

        [Fact]
        public void SeriesName_ReturnsChartSeriesName()
        {
            // Arrange
            IChart chart = (IChart)_fixture.Pre025.Slides[0].Shapes.First(sp => sp.Id == 5);

            // Act
            string seriesNameCase1 = chart.SeriesCollection[0].Name;
            string seriesNameCase2 = chart.SeriesCollection[2].Name;

            // Assert
            seriesNameCase1.Should().BeEquivalentTo("Ряд 1");
            seriesNameCase2.Should().BeEquivalentTo("Ряд 3");
        }

        [Fact]
        public void Type_ReturnsChartType()
        {
            // Arrange
            IChart chartCase1 = (IChart)_fixture.Pre021.Slides[1].Shapes.First(sp => sp.Id == 3);
            IChart chartCase2 = (IChart)_fixture.Pre021.Slides[2].Shapes.First(sp => sp.Id == 4);
            IChart chartCase3 = (IChart)_fixture.Pre013.Slides[0].Shapes.First(sp => sp.Id == 5);
            IChart chartCase4 = (IChart)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 7);

            // Act
            ChartType chartTypeCase1 = chartCase1.Type;
            ChartType chartTypeCase2 = chartCase2.Type;
            ChartType chartTypeCase3 = chartCase3.Type;
            ChartType chartTypeCase4 = chartCase4.Type;

            // Assert
            chartTypeCase1.Should().Be(ChartType.BubbleChart);
            chartTypeCase2.Should().Be(ChartType.ScatterChart);
            chartTypeCase3.Should().Be(ChartType.Combination);
            chartTypeCase4.Should().Be(ChartType.PieChart);
        }

        [Fact]
        public void GeometryType_GetterReturnsRectangle()
        {
            // Arrange
            IChart chart = (IChart)_fixture.Pre018.Slides[0].Shapes.First(sp => sp.Id == 6);

            // Act-Assert
            chart.GeometryType.Should().Be(GeometryType.Rectangle);
        }
    }
}
