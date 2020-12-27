using System.Linq;
using ShapeCrawler.Models;
using SlideDotNet.Models;
using Xunit;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.UnitTests
{
    public class TestFile_024
    {
        [Fact]
        public void Slide_Shapes_Test()
        {
            // Arrange
            var pre = new Presentation(Properties.Resources._024);
            var sld2 = pre.Slides[1];
            var chart = sld2.Shapes.First(x => x.Id == 5).Chart;

            // Act
            var hasXValues = chart.HasXValues;
            var XValue = chart.XValues[0];

            // Assert
            var shapes = pre.Slides[0].Shapes;
            Assert.NotNull(shapes);
            Assert.True(hasXValues);
            Assert.Equal(10, XValue);
        }
    }
}
