using System.Collections.Generic;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Enums;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using Xunit;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.UnitTests
{
    public class ShapeTests : IClassFixture<TestFileFixture>
    {
        private readonly TestFileFixture _fixture;

        public ShapeTests(TestFileFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void PlaceholderTypeXAndTextFrameTextProperties_ReturnCorrectValues()
        {
            // Arrange
            var pre021 = _fixture.Pre021;
            var sld4Shapes = pre021.Slides[3].Shapes;
            var sld3Shape2 = sld4Shapes.First(s => s.Id == 2);

            // Act
            PlaceholderType placeholderType = sld3Shape2.PlaceholderType;
            long xAxis = sld3Shape2.X;
            string text = sld3Shape2.TextFrame.Text;

            // Assert
            Assert.Equal(PlaceholderType.Footer, placeholderType);
            Assert.Equal(3653579, xAxis);
            Assert.Equal("test footer", text);
        }

        [Theory]
        [MemberData(nameof(ReturnsCorrectGeometryTypeValueTestCases))]
        public void GeometryType_ReturnsCorrectGeometryTypeValue(Shape shape, GeometryType expectedGeometryType)
        {
            // Assert
            shape.GeometryType.Should().BeEquivalentTo(expectedGeometryType);
        }

        public static IEnumerable<object[]> ReturnsCorrectGeometryTypeValueTestCases()
        {
            var pre021 = Presentation.Open(Properties.Resources._021);
            var shapes = pre021.Slides[3].Shapes;
            var shape2 = shapes.First(s => s.Id == 2);
            var shape3 = shapes.First(s => s.Id == 3);

            yield return new object[] { shape2, GeometryType.Rectangle };
            yield return new object[] { shape3, GeometryType.Ellipse };
        }
    }
}
