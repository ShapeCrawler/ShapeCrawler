using System.Linq;
using FluentAssertions;
using ShapeCrawler.Experiment;
using ShapeCrawler.SlideMaster;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;

namespace ShapeCrawler.Tests.Unit
{
    public class SlideMasterTests : IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public SlideMasterTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact(Skip = "In Progress")]
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

        [Fact(Skip = "In Progress")]
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

        [Fact(Skip = "In Progress")]
        public void MasterShapePlaceholderType_ReturnPlaceholderTypeOfTheMasterShape_WhenTheMasterShapeIsPlaceholder()
        {
            // Arrange
            //SlideMasterSc slideMaster = _fixture.Pre001.SlideMasters[0];
            //MasterShape masterAutoShapeCase1 = (MasterShape)slideMaster.Shapes.First(sp => sp.Id == 2);
            //MasterShape masterAutoShapeCase2 = (MasterShape)slideMaster.Shapes.First(sp => sp.Id == 8);
            //MasterShape masterAutoShapeCase3 = (MasterShape)slideMaster.Shapes.First(sp => sp.Id == 7);

            //// Act
            //PlaceholderType? shapePlaceholderTypeCase1 = masterAutoShapeCase1.PlaceholderType;
            //PlaceholderType? shapePlaceholderTypeCase2 = masterAutoShapeCase2.PlaceholderType;
            //PlaceholderType? shapePlaceholderTypeCase3 = masterAutoShapeCase3.PlaceholderType;

            //// Assert
            //shapePlaceholderTypeCase1.Should().Be(PlaceholderType.Title);
            //shapePlaceholderTypeCase2.Should().BeNull();
            //shapePlaceholderTypeCase3.Should().BeNull();
        }

        [Fact(Skip = "In Progress")]
        public void ShapeGeometryType_ReturnsShapesGeometryFormType()
        {
            // Arrange
            SlideMasterSc slideMaster = _fixture.Pre001.SlideMasters[0];
            BaseShape shapeCase1 = slideMaster.Shapes.First(sp => sp.Id == 2);
            BaseShape shapeCase2 = slideMaster.Shapes.First(sp => sp.Id == 8);

            // Act
            GeometryType geometryTypeCase1 = shapeCase1.GeometryType;
            GeometryType geometryTypeCase2 = shapeCase2.GeometryType;

            // Assert
            geometryTypeCase1.Should().Be(GeometryType.Rectangle);
            geometryTypeCase2.Should().Be(GeometryType.Custom);
        }

        [Fact(Skip = "In Progress")]
        public void MasterAutoShapeTextBoxText_ReturnsText_WhenTheAutoShapesTextBoxIsNotEmpty()
        {
            // Arrange
            SlideMasterSc slideMaster = _fixture.Pre001.SlideMasters[0];
            IAutoShape autoShape = (IAutoShape)slideMaster.Shapes.First(sp => sp.Id == 8);

            // Act-Assert
            autoShape.TextBox.Text.Should().BeEquivalentTo("id8");
        }

        [Fact(Skip = "In Progress")]
        public void MasterAutoShapeTextBoxParagraphPortionFontSize_ReturnsTextPortionFontSize()
        {
            // Arrange
            //SlideMasterSc slideMaster = _fixture.Pre001.SlideMasters[0];
            //MasterAutoShape masterAutoShape = (MasterAutoShape)slideMaster.Shapes.First(sp => sp.Id == 8);

            //// Act
            //int portionFontSize = masterAutoShape.TextBox.Paragraphs[0].Portions[0].Font.Size;

            //// Assert
            //portionFontSize.Should().Be(1800);
        }
    }
}
