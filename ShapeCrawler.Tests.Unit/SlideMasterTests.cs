using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using ShapeCrawler.Enums;
using ShapeCrawler.Models;
using ShapeCrawler.SlideMaster;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;

namespace ShapeCrawler.Tests.Unit
{
    [SuppressMessage("ReSharper", "SuggestVarOrType_BuiltInTypes")]
    public class SlideMasterTests : IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public SlideMasterTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void ShapeXAndY_ReturnXAndYAxesCoordinatesOfTheMasterShape()
        {
            // Arrange
            SlideMasterSc slideMaster = _fixture.Pre001.SlideMasters[0];
            BaseShape shape = slideMaster.Shapes.First(sp => sp.Id == 2);

            // Act
            long shapeXCoordinate = shape.X;
            long shapeYCoordinate = shape.Y;

            // Assert
            shapeXCoordinate.Should().Be(838200);
            shapeYCoordinate.Should().Be(365125);
        }

        [Fact]
        public void ShapeWidthAndHeight_ReturnWidthAndHeightSizesOfTheMaster()
        {
            // Arrange
            SlideMasterSc slideMaster = _fixture.Pre001.SlideMasters[0];
            BaseShape shape = slideMaster.Shapes.First(sp => sp.Id == 2);

            // Act
            long shapeWidth = shape.Width;
            long shapeHeight = shape.Height;

            // Assert
            shapeWidth.Should().Be(10515600);
            shapeHeight.Should().Be(1325563);
        }

        [Fact]
        public void MasterAutoShapePlaceholderType_ReturnPlaceholderTypeOfTheMasterShape_WhenTheMasterShapeIsPlaceholder()
        {
            // Arrange
            SlideMasterSc slideMaster = _fixture.Pre001.SlideMasters[0];
            BaseShape shape = slideMaster.Shapes.First(sp => sp.Id == 2);
            MasterAutoShape masterAutoShape = (MasterAutoShape) shape;

            // Act
            PlaceholderType? shapePlaceholderType = masterAutoShape.PlaceholderType;

            // Assert
            shapePlaceholderType.Should().Be(PlaceholderType.Title);
        }

        [Fact(Skip = "The feature is in progress")]
        public void ShapeGeometryType_ReturnsShapesGeometryFormType()
        {
            // Arrange
            SlideMasterSc slideMaster = _fixture.Pre001.SlideMasters[0];
            BaseShape shape = slideMaster.Shapes.First(sp => sp.Id == 2);

            // Act
            GeometryType shapeGeometryType = shape.GeometryType;

            // Assert
            shapeGeometryType.Should().Be(GeometryType.Rectangle);
        }
    }
}
