using System.IO;
using System.Linq;
using FluentAssertions;
using Xunit;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.UnitTests
{
    /// <summary>
    /// Represents a class whose instance is created for each test.
    /// </summary>
    public class TestFile_025 : IClassFixture<Presentation25Fixture>
    {
        private readonly Presentation25Fixture _pre25Fixture;

        public TestFile_025(Presentation25Fixture pre25Fixture)
        {
            _pre25Fixture = pre25Fixture;
        }

        [Fact]
        public void Chart_Test()
        {
            // Arrange
            var sld1 = _pre25Fixture.Presentation.Slides[0];
            var sld2 = _pre25Fixture.Presentation.Slides[1];
            var chart8 = sld1.Shapes.First(x => x.Id == 8).Chart;
            var chart4 = sld1.Shapes.First(x => x.Id == 4).Chart;
            var chart5 = sld1.Shapes.First(x => x.Id == 5).Chart;
            var chart11 = sld2.Shapes.First(x => x.Id == 11).Chart;
            var chart4ChildCat = chart4.Categories[0];
            var chart5SeriesCollection = chart5.SeriesCollection;

            // Act
            var chart8HasXValues = chart8.HasXValues;
            var chart11HasXValues = chart11.HasXValues;
            var chart4ChildCatVal = chart4ChildCat.Name;
            var chart4ParentCatVal = chart4ChildCat.Parent.Name;
            var serName1 = chart5SeriesCollection[0].Name;
            var serName3 = chart5SeriesCollection[2].Name;

            // Assert
            Assert.False(chart8HasXValues);
            Assert.False(chart11HasXValues);
            Assert.Equal("Dresses", chart4ChildCatVal);
            Assert.Equal("Clothing", chart4ParentCatVal);
            Assert.Equal("Ряд 1", serName1);
            serName3.Should().Be("Ряд 3");
        }

        [Fact]
        public void Chart_Test_2()
        {
            // Arrange
            var sld1 = _pre25Fixture.Presentation.Slides[0];
            var sld2 = _pre25Fixture.Presentation.Slides[1];
            var chart4 = sld2.Shapes.First(x => x.Id == 4).Chart;
            var chart7 = sld1.Shapes.First(x => x.Id == 7).Chart;

            // Act
            var pointValue = chart4.SeriesCollection[0].PointValues[0];
            var chartTitle = chart7.Title;

            // Assert
            Assert.True(pointValue > 0);
            Assert.NotNull(chartTitle);
        }

        [Fact]
        public void SaveScheme_Test()
        {
            // Arrange
            var sld3 = _pre25Fixture.Presentation.Slides[2];
            var chart3 = sld3.Shapes.First(x => x.Id == 7);
            var stream = new MemoryStream();

            // Act
            sld3.SaveScheme(stream);
            var chartX = chart3.X;

            // Assert
            Assert.True(stream.Length > 0);
            Assert.Equal(757383, chartX);
        }
    }
}
