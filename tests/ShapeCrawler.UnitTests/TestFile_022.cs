using System.Linq;
using ShapeCrawler.Enums;
using ShapeCrawler.Models;
using SlideDotNet.Models;
using Xunit;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.UnitTests
{
    public class TestFile_022
    {
        [Fact]
        public void Slide_Shapes_Test()
        {
            // Arrange
            var pre = new Presentation(Properties.Resources._022);
            var spPic3 = pre.Slides[0].Shapes.First(x => x.Id == 3);

            // Act
            var shapes = pre.Slides[0].Shapes; // act-assert
            var geometry = spPic3.GeometryType;

            // Arrange
            Assert.Equal(GeometryType.Ellipse, geometry);
        }
    }
}
