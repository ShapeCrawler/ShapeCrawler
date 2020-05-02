using System.Linq;
using SlideDotNet.Enums;
using SlideDotNet.Models;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace SlideDotNet.Tests
{
    public class TestFile_021
    {
        [Fact]
        public void Chart_Test()
        {
            // Arrange
            var pre = new PresentationEx(Properties.Resources._021);
            var shapes1 = pre.Slides[0].Shapes;
            var shapes2 = pre.Slides[1].Shapes;
            var sp108 = shapes1.Single(x => x.Id == 108);
            var chart3 = shapes1.Single(x => x.Id == 3).Chart;
            var sld1Chart4 = shapes1.Single(x => x.Id == 4).Chart;
            var sld2Chart4 = shapes2.First(x => x.Id == 4).Chart;
            var lineChartSeries = sld2Chart4.SeriesCollection[1];

            // Act
            var fill = sp108.Fill; //assert: do not throw exception
            
            var chartTypeBar = chart3.SeriesCollection[1].Type;
            var pValueBar = chart3.SeriesCollection[1].PointValues[0];
            var chartTypeScatter = chart3.SeriesCollection[2].Type;
            var pValueScatter = chart3.SeriesCollection[2].PointValues[0];
            var category = sld1Chart4.Categories[0];
            var pv = lineChartSeries.PointValues[0];

            // Assert
            Assert.Equal(ChartType.BarChart, chartTypeBar);
            Assert.Equal(56, pValueBar);
            Assert.Equal(ChartType.ScatterChart, chartTypeScatter);
            Assert.Equal(44, pValueScatter);
            Assert.Equal("2015", category.Value);
            Assert.Equal(17.35, pv);
        }

        [Fact]
        public void Chart3_Test()
        {
            // Arrange
            var pre = new PresentationEx(Properties.Resources._021);
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
            var pre = new PresentationEx(Properties.Resources._021);
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

        [Fact]
        public void Footer_Test()
        {
            // Arrange
            var pre = new PresentationEx(Properties.Resources._021);
            var sld4Shapes = pre.Slides[3].Shapes;

            // Act
            var footerShape = sld4Shapes.First(s => s.Id == 2);
            var ellipse = sld4Shapes.First(s => s.Id == 3).GeometryType;
            var type = footerShape.PlaceholderType;
            var text = footerShape.TextFrame.Text;
            var x = footerShape.X;
            var geometry = footerShape.GeometryType;

            // Assert
            Assert.Equal(PlaceholderType.Footer, type);
            Assert.Equal("test footer", text);
            Assert.Equal(3653579, x);
            Assert.Equal(GeometryType.Rectangle, geometry);
            Assert.Equal(GeometryType.Ellipse, ellipse);
        }
    }
}
