using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Enums;
using ShapeCrawler.Extensions;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.UnitTests.Helpers;
using ShapeCrawler.UnitTests.Properties;
using Xunit;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.UnitTests
{
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    public class ShapeTests : IClassFixture<ReadOnlyTestPresentations>
    {
        private readonly ReadOnlyTestPresentations _fixture;

        public ShapeTests(ReadOnlyTestPresentations fixture)
        {
            _fixture = fixture;
        }

        [Theory]
        [MemberData(nameof(PlaceholderTypePropertyTestCases))]
        public void PlaceholderType_GetterReturnsPlaceholderTypeOfTheShape(Shape shape, PlaceholderType expectedType)
        {
            // Act
            var actualType = shape.PlaceholderType;

            // Assert
            actualType.Should().Be(expectedType);
        }

        public static IEnumerable<object[]> PlaceholderTypePropertyTestCases()
        {
            var shape = Presentation.Open(Resources._021, false).
                                                    Slides[3].
                                                    Shapes.First(sp => sp.Id == 2);
            yield return new object[] { shape, PlaceholderType.Footer };

            shape = Presentation.Open(Resources._008, false)
                .Slides[0]
                .Shapes.First(sp => sp.Id == 3);
            yield return new object[] { shape, PlaceholderType.DateAndTime };
        }

        [Fact]
        public void FillType_GetterReturnsFillTypeByWhichTheShapeIsFilled()
        {
            // Arrange
            Shape shapeCase1 = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 4);
            Shape shapeCase2 = _fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 2);

            // Act
            FillType shapeFillTypeCase1 = shapeCase1.Fill.Type;
            FillType shapeFillTypeCase2 = shapeCase2.Fill.Type;

            // Assert
            shapeFillTypeCase1.Should().Be(FillType.Picture);
            shapeFillTypeCase2.Should().Be(FillType.Solid);
        }

        [Fact]
        public void FillSolidColorName_ReturnsSolidColorNameByWhichTheShapeIsFilled()
        {
            // Arrange
            Shape shape = _fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 2);

            // Act
            var shapeSolidColorName = shape.Fill.SolidColor.Name;

            // Assert
            shapeSolidColorName.Should().BeEquivalentTo("ff0000");
        }

        [Fact]
        public void Fill_ReturnsNull_WhenShapeIsNotFilled()
        {
            // Arrange
            Shape shape = _fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 6);

            // Act
            Fill shapeFill = shape.Fill;

            // Assert
            shapeFill.Should().BeNull();
        }

        [Fact]
        public void Y_GetterReturnsYCoordinateOfTheShape()
        {
            // Arrange
            var shape = _fixture.Pre006.Slides.First().Shapes.First(sp => sp.Id == 2);

            // Act
            var yCoordinate = shape.Y;

            // Assert
            yCoordinate.Should().Be(1122363);
        }

        [Fact]
        public void X_GetterReturnsXCoordinateOfTheShape()
        {
            // Arrange
            var shapeCase1 = _fixture.Pre021.Slides[3].Shapes.First(sp => sp.Id == 2);
            var shapeCase2 = _fixture.Pre008.Slides[0].Shapes.First(sp => sp.Id == 3);
            var shapeCase3 = _fixture.Pre006.Slides[0].Shapes.First(sp => sp.Id == 2);
            var shapeCase4 = _fixture.Pre009.Slides[1].Shapes.Single(sp => sp.ContentType.Equals(ShapeContentType.Group)).
                                GroupedShapes.Single(sp => sp.Id.Equals(5));
            // Act
            var xCoordinateCase1 = shapeCase1.X;
            var xCoordinateCase2 = shapeCase2.X;
            var xCoordinateCase3 = shapeCase3.X;
            var xCoordinateCase4 = shapeCase4.X;

            // Assert
            xCoordinateCase1.Should().Be(3653579);
            xCoordinateCase2.Should().Be(628650);
            xCoordinateCase3.Should().Be(1524000);
            xCoordinateCase4.Should().Be(1581846);
        }

        [Fact]
        public void XAndWidth_SettersSetXAndWidthOfTheShape()
        {
            // Arrange
            var presentation = Presentation.Open(Resources._006_1_slides, true);
            var shape = presentation.Slides.First().Shapes.First(sp => sp.Id == 3);
            var stream = new MemoryStream();
            const int newX = 4000000;
            const int newWidth = 6000000;

            // Act
            shape.X = newX;
            shape.Width = newWidth;
            presentation.SaveAs(stream);

            // Assert
            presentation = Presentation.Open(stream, false);
            shape = presentation.Slides.First().Shapes.First(sp => sp.Id == 3);
            shape.X.Should().Be(newX);
            shape.Width.Should().Be(newWidth);
        }

        [Fact]
        public void WidthAndHeight_ReturnWidthAndHeightSizesOfTheShape()
        {
            // Arrange
            var shapeCase1 = _fixture.Pre006.Slides[0].Shapes.First(sp => sp.Id == 2);
            var shapeCase2 = _fixture.Pre009.Slides[1].Shapes.Single(sp => sp.ContentType.Equals(ShapeContentType.Group)).
                                GroupedShapes.Single(sp => sp.Id.Equals(5));

            // Act
            var shapeWidthCase1 = shapeCase1.Width;
            var shapeHeightCase1 = shapeCase1.Height;
            var shapeWidthCase2 = shapeCase2.Width;
            var shapeHeightCase2 = shapeCase2.Height;

            // Assert
            shapeWidthCase1.Should().Be(9144000);
            shapeHeightCase1.Should().Be(1425528);
            shapeWidthCase2.Should().Be(1181377);
            shapeHeightCase2.Should().Be(654096);
        }

        [Theory]
        [MemberData(nameof(ReturnsCorrectGeometryTypeValueTestCases))]
        public void GeometryType_ReturnsCorrectGeometryTypeValue(Shape shape, GeometryType expectedGeometryType)
        {
            // Assert
            shape.GeometryType.Should().BeEquivalentTo(expectedGeometryType);
        }

        [Fact]
        public void IsPlaceholderAndIsGrouped_ReturnFalseValues_WhenTheShapeIsNotAPlaceholderAndItIsNotGrouped()
        {
            // Arrange
            var shape = _fixture.Pre006.Slides[0].Shapes.First(x => x.Id == 3);

            // Act
            var isPlaceholderShape = shape.IsPlaceholder;
            var isGroupedShape = shape.IsGrouped;

            // Assert
            isPlaceholderShape.Should().BeFalse();
            isGroupedShape.Should().BeFalse();
        }

        [Fact]
        public void CustomData_ReturnNull_WhenShapeHasNotCustomData()
        {
            // Arrange
            var shape = _fixture.Pre009.Slides.First().Shapes.First();

            // Act
            var shapeCustomData = shape.CustomData;

            // Assert
            shapeCustomData.Should().BeNull();
        }


        [Fact]
        public void CustomData_ReturnsCustomDataOfTheShape_WhenShapeWasAssignedSomeCustomData()
        {
            // Arrange
            const string customDataString = "Test custom data";
            var savedPreStream = new MemoryStream();
            var presentation = Presentation.Open(Resources._009, true);
            var shape = presentation.Slides.First().Shapes.First();

            // Act
            shape.CustomData = customDataString;
            presentation.SaveAs(savedPreStream);

            // Assert
            presentation = Presentation.Open(savedPreStream, false);
            shape = presentation.Slides.First().Shapes.First();
            shape.CustomData.Should().Be(customDataString);
        }

        #region Helpers

        public static IEnumerable<object[]> ReturnsCorrectGeometryTypeValueTestCases()
        {
            var pre021 = Presentation.Open(Resources._021);
            var shapes = pre021.Slides[3].Shapes;
            var shape2 = shapes.First(s => s.Id == 2);
            var shape3 = shapes.First(s => s.Id == 3);

            yield return new object[] { shape2, GeometryType.Rectangle };
            yield return new object[] { shape3, GeometryType.Ellipse };
        }

        #endregion Helpers
    }
}
