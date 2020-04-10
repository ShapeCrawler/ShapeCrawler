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
    public class TestFile_022
    {
        [Fact]
        public void Slide_Shapes_Test()
        {
            // Arrange
            var pre = new PresentationEx(Properties.Resources._022);
            var spPic3 = pre.Slides[0].Shapes.First(x => x.Id == 3);

            // Act
            var shapes = pre.Slides[0].Shapes; // act-assert
            var geometry = spPic3.GeometryType;

            // Arrange
            Assert.Equal(GeometryType.Ellipse, geometry);
        }
    }
}
