using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using FluentAssertions;
using ShapeCrawler.Enums;
using ShapeCrawler.Models.Styles;
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
    public class ShapeTests : IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public ShapeTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Theory]
        [MemberData(nameof(TestCasesPlaceholderType))]
        public void PlaceholderType_GetterReturnsPlaceholderTypeOfTheShape(ShapeSc shape, PlaceholderType expectedType)
        {
            // Act
            PlaceholderType? actualType = shape.PlaceholderType;

            // Assert
            actualType.Should().Be(expectedType);
        }

        public static IEnumerable<object[]> TestCasesPlaceholderType()
        {
            ShapeSc shape = PresentationSc.Open(Resources._021, false).Slides[3].Shapes.First(sp => sp.Id == 2);
            yield return new object[] { shape, PlaceholderType.Footer };

            shape = PresentationSc.Open(Resources._008, false).Slides[0].Shapes.First(sp => sp.Id == 3);
            yield return new object[] { shape, PlaceholderType.DateAndTime };

            shape = PresentationSc.Open(Resources._019, false).Slides[0].Shapes.First(sp => sp.Id == 2);
            yield return new object[] { shape, PlaceholderType.SlideNumber };

            shape = PresentationSc.Open(Resources._013, false).Slides[0].Shapes.First(sp => sp.Id == 281);
            yield return new object[] { shape, PlaceholderType.Custom };
        }

        [Fact]
        public void Fill_ReturnsNull_WhenShapeIsNotFilled()
        {
            // Arrange
            ShapeSc shapeEx = _fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 6);

            // Act
            ShapeFill shapeFill = shapeEx.Fill;

            // Assert
            shapeFill.Should().BeNull();
        }

        [Fact]
        public void FillType_GetterReturnsFillTypeByWhichTheShapeIsFilled()
        {
            // Arrange
            ShapeSc shapeExCase1 = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 4);
            ShapeSc shapeExCase2 = _fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 2);

            // Act
            FillType shapeFillTypeCase1 = shapeExCase1.Fill.Type;
            FillType shapeFillTypeCase2 = shapeExCase2.Fill.Type;

            // Assert
            shapeFillTypeCase1.Should().Be(FillType.Picture);
            shapeFillTypeCase2.Should().Be(FillType.Solid);
        }

        [Fact]
        public void FillSolidColorName_ReturnsSolidColorNameByWhichTheShapeIsFilled()
        {
            // Arrange
            ShapeSc shapeEx = _fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 2);

            // Act
            var shapeSolidColorName = shapeEx.Fill.SolidColor.Name;

            // Assert
            shapeSolidColorName.Should().BeEquivalentTo("ff0000");
        }


        [Fact]
        public async void FillPictureGetImageBytes_ReturnsImageByWhichTheShapeIsFilled()
        {
            // Arrange
            ShapeSc shapeEx = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 4);

            // Act
            var shapeFilledImage = await shapeEx.Fill.Picture.GetImageBytes();

            // Assert
            shapeFilledImage.Length.Should().BePositive();
        }

        [Fact]
        public async void FillPictureSetImage_MethodSetsImageForPictureFilledShape()
        {
            // Arrange
            var presentation = PresentationSc.Open(Resources._009, true);
            ShapeSc shapeEx = presentation.Slides[2].Shapes.First(sp => sp.Id.Equals(4));
            var newImage = new MemoryStream(Resources.test_image_2);
            var imageSizeBefore = (await shapeEx.Fill.Picture.GetImageBytes()).Length;

            // Act
            shapeEx.Fill.Picture.SetImage(newImage);

            // Assert
            var imageSizeAfter = (await shapeEx.Fill.Picture.GetImageBytes()).Length;
            imageSizeAfter.Should().NotBe(imageSizeBefore);
        }

        [Fact]
        public void XAndY_ReturnXAndYAxesShapeCoordinatesOnTheSlide()
        {
            // Arrange
            ShapeSc shapeExCase1 = _fixture.Pre021.Slides[3].Shapes.First(sp => sp.Id == 2);
            ShapeSc shapeExCase2 = _fixture.Pre008.Slides[0].Shapes.First(sp => sp.Id == 3);
            ShapeSc shapeExCase3 = _fixture.Pre006.Slides[0].Shapes.First(sp => sp.Id == 2);
            ShapeSc shapeExCase4 = _fixture.Pre009.Slides[1].Shapes.Single(sp => sp.ContentType.Equals(ShapeContentType.Group)).
                                GroupedShapes.Single(sp => sp.Id.Equals(5));
            ShapeSc shapeExCase5 = _fixture.Pre018.Slides[0].Shapes.First(sp => sp.Id == 7);
            ShapeSc shapeExCase6 = _fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 9);
            ShapeSc shapeExCase7 = _fixture.Pre025.Slides[2].Shapes.First(sp => sp.Id == 7);

            // Act
            long xCoordinateCase1 = shapeExCase1.X;
            long xCoordinateCase2 = shapeExCase2.X;
            long xCoordinateCase3 = shapeExCase3.X;
            long xCoordinateCase4 = shapeExCase4.X;
            long xCoordinateCase6 = shapeExCase6.X;
            long xCoordinateCase7 = shapeExCase7.X;
            long yCoordinateCase3 = shapeExCase3.Y;
            long yCoordinateCase5 = shapeExCase5.Y;
            long yCoordinateCase6 = shapeExCase6.Y;

            // Assert
            xCoordinateCase1.Should().Be(3653579);
            xCoordinateCase2.Should().Be(628650);
            xCoordinateCase3.Should().Be(1524000);
            xCoordinateCase4.Should().Be(1581846);
            xCoordinateCase6.Should().Be(699323);
            xCoordinateCase7.Should().Be(757383);
            yCoordinateCase3.Should().Be(1122363);
            yCoordinateCase5.Should().Be(4);
            yCoordinateCase6.Should().Be(3463288);
        }

        [Fact]
        public void XAndWidth_SettersSetXAndWidthOfTheShape()
        {
            // Arrange
            var presentation = PresentationSc.Open(Resources._006_1_slides, true);
            var shape = presentation.Slides.First().Shapes.First(sp => sp.Id == 3);
            var stream = new MemoryStream();
            const int newX = 4000000;
            const int newWidth = 6000000;

            // Act
            shape.X = newX;
            shape.Width = newWidth;
            presentation.SaveAs(stream);

            // Assert
            presentation = PresentationSc.Open(stream, false);
            shape = presentation.Slides.First().Shapes.First(sp => sp.Id == 3);
            shape.X.Should().Be(newX);
            shape.Width.Should().Be(newWidth);
        }

        [Fact]
        public void WidthAndHeight_ReturnWidthAndHeightSizesOfTheShape()
        {
            // Arrange
            ShapeSc shapeCase1 = _fixture.Pre006.Slides[0].Shapes.First(sp => sp.Id == 2);
            ShapeSc shapeCase2 = _fixture.Pre009.Slides[1].Shapes.Single(sp => sp.ContentType.Equals(ShapeContentType.Group)).
                                    GroupedShapes.Single(sp => sp.Id.Equals(5));
            ShapeSc shapeCase3 = _fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 9);

            // Act
            long shapeWidthCase1 = shapeCase1.Width;
            long shapeWidthCase2 = shapeCase2.Width;
            long shapeWidthCase3 = shapeCase3.Width;
            long shapeHeightCase1 = shapeCase1.Height;
            long shapeHeightCase2 = shapeCase2.Height;
            long shapeHeightCase3 = shapeCase3.Height;

            // Assert
            shapeWidthCase1.Should().Be(9144000);
            shapeWidthCase2.Should().Be(1181377);
            shapeWidthCase3.Should().Be(485775);
            shapeHeightCase1.Should().Be(1425528);
            shapeHeightCase2.Should().Be(654096);
            shapeHeightCase3.Should().Be(373062);
        }

        [Theory]
        [MemberData(nameof(GeometryTypeTestCases))]
        public void GeometryType_ReturnsCorrectGeometryTypeValue(ShapeSc shapeEx, GeometryType expectedGeometryType)
        {
            // Assert
            shapeEx.GeometryType.Should().BeEquivalentTo(expectedGeometryType);
        }

        [Fact]
        public void ContentTypes_ReturnOLEObjectTypeEnum_WhenShapeContainsOLEObjectData()
        {
            // Arrange
            ShapeSc shape = _fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 8);

            // Act
            ShapeContentType shapeContentType = shape.ContentType;

            // Assert
            shapeContentType.Should().Be(ShapeContentType.OLEObject);
        }

        public static IEnumerable<object[]> GeometryTypeTestCases()
        {
            var pre021 = PresentationSc.Open(Resources._021,false);
            var shapes = pre021.Slides[3].Shapes;
            var shapeCase1 = shapes.First(sp => sp.Id == 2);
            var shapeCase2 = shapes.First(sp => sp.Id == 3);

            yield return new object[] { shapeCase1, GeometryType.Rectangle };
            yield return new object[] { shapeCase2, GeometryType.Ellipse };
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
        public void CustomData_ReturnsNull_WhenShapeHasNotCustomData()
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
            var presentation = PresentationSc.Open(Resources._009, true);
            var shape = presentation.Slides.First().Shapes.First();

            // Act
            shape.CustomData = customDataString;
            presentation.SaveAs(savedPreStream);

            // Assert
            presentation = PresentationSc.Open(savedPreStream, false);
            shape = presentation.Slides.First().Shapes.First();
            shape.CustomData.Should().Be(customDataString);
        }

        [Fact]
        public void Name_ReturnsShapeNameString()
        {
            // Arrange
            ShapeSc shape = _fixture.Pre009.Slides[1].Shapes.First(sp => sp.Id == 8);

            // Act
            string shapeName = shape.Name;

            // Assert
            shapeName.Should().BeEquivalentTo("Object 2");
        }

        [Fact]
        public void HasTextFrame_ReturnsFalse_WhenTheShapeDoesNotContainText()
        {
            // Arrange
            ShapeSc shape = _fixture.Pre009.Slides[4].Shapes.First(sp => sp.Id == 5);

            // Act
            bool hasTextFrame = shape.HasTextBox;

            // Assert
            hasTextFrame.Should().BeFalse();
        }
    }
}
