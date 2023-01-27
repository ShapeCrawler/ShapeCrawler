#if DEBUG

using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Media;
using ShapeCrawler.Shapes;
using ShapeCrawler.UnitTests.Helpers;
using ShapeCrawler.UnitTests.Helpers.Attributes;
using Xunit;
using TestHelper = ShapeCrawler.Tests.Shared.TestHelper;

namespace ShapeCrawler.UnitTests
{
    public class ShapeTests : ShapeCrawlerTest
    {
        [Theory]
        [SlideShapeData("021.pptx", 4, 2, SCPlaceholderType.Footer)]
        [SlideShapeData("008.pptx", 1, 3, SCPlaceholderType.DateAndTime)]
        [SlideShapeData("019.pptx", 1, 2, SCPlaceholderType.SlideNumber)]
        [SlideShapeData("013.pptx", 1, 281, SCPlaceholderType.Content)]
        [SlideShapeData("autoshape-case016.pptx", 1, "Content Placeholder 1", SCPlaceholderType.Content)]
        [SlideShapeData("autoshape-case016.pptx", 1, "Text Placeholder 1", SCPlaceholderType.Text)]
        [SlideShapeData("autoshape-case016.pptx", 1, "Picture Placeholder 1", SCPlaceholderType.Picture)]
        [SlideShapeData("autoshape-case016.pptx", 1, "Table Placeholder 1", SCPlaceholderType.Table)]
        [SlideShapeData("autoshape-case016.pptx", 1, "SmartArt Placeholder 1", SCPlaceholderType.SmartArt)]
        [SlideShapeData("autoshape-case016.pptx", 1, "Media Placeholder 1", SCPlaceholderType.Media)]
        [SlideShapeData("autoshape-case016.pptx", 1, "Online Image Placeholder 1", SCPlaceholderType.OnlineImage)]
        public void PlaceholderType_returns_placeholder_type(IShape shape, SCPlaceholderType expectedType)
        {
            // Act
            var placeholderType = shape.Placeholder!.Type;

            // Assert
            placeholderType.Should().Be(expectedType);
        }

        [Fact]
        public void AudioShape_BinaryData_returns_audio_bytes()
        {
            // Arrange
            var pptxStream = GetTestStream("audio-case001.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var audioShape = pres.Slides[0].Shapes.GetByName<IAudioShape>("Audio 1");

            // Act
            var bytes = audioShape.BinaryData;

            // Assert
            bytes.Should().NotBeEmpty();
        }

        [Fact]
        public void AudioShape_MIME_returns_MIME_type_of_audio_content()
        {
            // Arrange
            var pptxStream = GetTestStream("audio-case001.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var audioShape = pres.Slides[0].Shapes.GetByName<IAudioShape>("Audio 1");

            // Act
            var mime = audioShape.MIME;

            // Assert
            mime.Should().Be("audio/mpeg");
        }

        [Fact(Skip = "On Hold ")]
        public void Duplicate_duplicates_AutoShape()
        {
            // Arrange
            var pptx = TestHelper.GetStream("autoshape-case015.pptx");
            var pres = SCPresentation.Open(pptx);
            var shape = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 6");

            // Act
            IAutoShape shapeCopy = shape.Duplicate();

            // Assert
            shapeCopy.X.Should().Be(shape.X);
            shapeCopy.Width.Should().Be(shape.Width);
            shapeCopy.TextFrame.Text.Should().Be(shapeCopy.TextFrame.Text);
        }
        
        [Fact]
        public void VideoShape_BinaryData_returns_video_bytes()
        {
            // Arrange
            var pptxStream = GetTestStream("video-case001.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var videoShape = pres.Slides[0].Shapes.GetByName<IVideoShape>("Video 1");

            // Act
            var bytes = videoShape.BinaryData;

            // Assert
            bytes.Should().NotBeEmpty();
        }

        [Fact]
        public void AudioShape_MIME_returns_MIME_type_of_video_content()
        {
            // Arrange
            var pptxStream = GetTestStream("video-case001.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var videoShape = pres.Slides[0].Shapes.GetByName<IVideoShape>("Video 1");

            // Act
            var mime = videoShape.MIME;

            // Assert
            mime.Should().Be("video/mp4");
        }

        [Fact]
        public void PictureSetImage_ShouldNotImpactOtherPictureImage_WhenItsOriginImageIsShared()
        {
            // Arrange
            var pptx = GetTestStream("009_table.pptx");
            IPresentation presentation = SCPresentation.Open(pptx);
            IPicture picture5 = (IPicture)presentation.Slides[3].Shapes.First(sp => sp.Id == 5);
            IPicture picture6 = (IPicture)presentation.Slides[3].Shapes.First(sp => sp.Id == 6);
            int pic6LengthBefore = picture6.Image.BinaryData.GetAwaiter().GetResult().Length;
            MemoryStream modifiedPresentation = new();

            // Act
            picture5.Image.SetImage(TestFiles.Images.imageByteArray02);

            // Assert
            int pic6LengthAfter = picture6.Image.BinaryData.GetAwaiter().GetResult().Length;
            pic6LengthAfter.Should().Be(pic6LengthBefore);

            presentation.SaveAs(modifiedPresentation);
            presentation = SCPresentation.Open(modifiedPresentation);
            picture6 = (IPicture)presentation.Slides[3].Shapes.First(sp => sp.Id == 6);
            pic6LengthBefore = picture6.Image.BinaryData.GetAwaiter().GetResult().Length;
            pic6LengthAfter.Should().Be(pic6LengthBefore);
        }

        [Theory]
        [MemberData(nameof(TestCasesXGetter))]
        public void X_Getter_returns_x_coordinate_in_pixels(TestCase<IShape, int> testCase)
        {
            // Arrange
            var shape = testCase.Param1;
            var expectedX = testCase.Param2;
            
            // Act
            var xCoordinate = shape.X;
            
            // Assert
            xCoordinate.Should().Be(expectedX);
        }

        public static IEnumerable<object[]> TestCasesXGetter
        {
            get
            {
                var pptxStream1 = GetTestStream("021.pptx");
                var pres1 = SCPresentation.Open(pptxStream1);
                var shape1 = pres1.Slides[3].Shapes.GetById<IShape>(2);
                var testCase1 = new TestCase<IShape, int>(1, shape1, 383);
                yield return new object[] { testCase1 };
                
                var pptxStream2 = GetTestStream("008.pptx");
                var pres2 = SCPresentation.Open(pptxStream2);
                var shape2 = pres2.Slides[0].Shapes.GetById<IShape>(3);
                var testCase2 = new TestCase<IShape, int>(2, shape2, 66);
                yield return new object[] { testCase2 };
                
                var pptxStream3 = GetTestStream("006_1 slides.pptx");
                var pres3 = SCPresentation.Open(pptxStream3);
                var shape3 = pres3.Slides[0].Shapes.GetById<IShape>(2);
                var testCase3 = new TestCase<IShape, int>(3, shape3, 160);
                yield return new object[] { testCase3 };
                
                var pptxStream4 = GetTestStream("009_table.pptx");
                var pres4 = SCPresentation.Open(pptxStream4);
                var shape4 = pres4.Slides[1].Shapes.GetById<IShape>(9);
                var testCase4 = new TestCase<IShape, int>(4, shape4, 73);
                yield return new object[] { testCase4 };
                
                var pptxStream5 = GetTestStream("025_chart.pptx");
                var pres5 = SCPresentation.Open(pptxStream5);
                var shape5 = pres5.Slides[2].Shapes.GetById<IShape>(7);
                var testCase5 = new TestCase<IShape, int>(5, shape5, 79);
                yield return new object[] { testCase5 };
                
                var pptxStream6 = GetTestStream("018.pptx");
                var pres6 = SCPresentation.Open(pptxStream6);
                var shape6 = pres6.Slides[0].Shapes.GetByName<IShape>("Picture Placeholder 1");
                var testCase6 = new TestCase<IShape, int>(6, shape6, 9);
                yield return new object[] { testCase6 };
                
                var pptxStream7 = GetTestStream("009_table.pptx");
                var pres7 = SCPresentation.Open(pptxStream7);
                var shape7 = pres7.Slides[1].Shapes.GetById<IGroupShape>(7).Shapes.GetById<IShape>(5);
                var testCase7 = new TestCase<IShape, int>(7, shape7, 166);
                yield return new object[] { testCase7 };
            }
        }

        [Fact]
        public void Y_Getter_returns_y_coordinate_in_pixels()
        {
            // Arrange
            IShape shapeCase1 = SCPresentation.Open(GetTestStream("006_1 slides.pptx")).Slides[0].Shapes.First(sp => sp.Id == 2);
            IShape shapeCase2 = SCPresentation.Open(GetTestStream("018.pptx")).Slides[0].Shapes.First(sp => sp.Id == 7);
            IShape shapeCase3 = SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 9);
            float verticalResoulution = Helpers.TestHelper.VerticalResolution;

            // Act
            int yCoordinate1 = shapeCase1.Y;
            int yCoordinate2 = shapeCase2.Y;
            int yCoordinate3 = shapeCase3.Y;

            // Assert
            yCoordinate1.Should().Be((int)(1122363 * verticalResoulution / 914400));
            yCoordinate2.Should().Be((int)(4 * verticalResoulution / 914400));
            yCoordinate3.Should().Be((int)(3463288 * verticalResoulution / 914400));
        }

        [Fact]
        public void Id_returns_id()
        {
            // Arrange
            var pptxStream = GetTestStream("010.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var shape = pres.SlideMasters[0].Shapes.GetByName<IShape>("Date Placeholder 3");
            
            // Act
            var id = shape.Id;
            
            // Assert
            id.Should().Be(9);
        }

        [Theory]
        [SlideShapeData("001.pptx", 1, "TextBox 3")]
        [SlideShapeData("001.pptx", 1, "Head 1")]
        [SlideShapeData("autoshape-case015.pptx", 1, "Group 1")]
        public void Y_Setter_sets_y_coordinate(IShape shape)
        {
            // Act
            shape.Y = 100;

            // Assert
            shape.Y.Should().Be(100);
            var errors = PptxValidator.Validate(shape.SlideObject.Presentation);
            errors.Should().BeEmpty();
        }

        [Theory]
        [SlideShapeData("006_1 slides.pptx", 1, "Shape 1")]
        [SlideShapeData("001.pptx", 1, "Head 1")]
        [SlideShapeData("autoshape-case015.pptx", 1, "Group 1")]
        [SlideShapeData("table-case001.pptx", 1, "Table 1")]
        public void X_Setter_sets_x_coordinate(IShape shape)
        {
            // Arrange
            var pres = shape.SlideObject.Presentation;
            var slideIndex = shape.SlideObject.Number - 1;
            var shapeName = shape.Name;
            var stream = new MemoryStream();

            // Act
            shape.X = 400;

            // Assert
            pres.SaveAs(stream);
            pres = SCPresentation.Open(stream);
            shape = pres.Slides[slideIndex].Shapes.GetByName<IShape>(shapeName);
            shape.X.Should().Be(400);
            var errors = PptxValidator.Validate(shape.SlideObject.Presentation);
            errors.Should().BeEmpty();
        }
        
        [Theory]
        [SlideShapeData("006_1 slides.pptx", 1, "Shape 1")]
        [SlideShapeData("autoshape-case015.pptx", 1, "Group 1")]
        public void Width_Setter_sets_width(IShape shape)
        {
            // Arrange
            var pres = shape.SlideObject.Presentation;
            var slideIndex = shape.SlideObject.Number - 1;
            var shapeName = shape.Name;
            var stream = new MemoryStream();

            // Act
            shape.Width = 600;

            // Assert
            pres.SaveAs(stream);
            pres = SCPresentation.Open(stream);
            shape = pres.Slides[slideIndex].Shapes.GetByName<IShape>(shapeName);
            shape.Width.Should().Be(600);
            var errors = PptxValidator.Validate(shape.SlideObject.Presentation);
            errors.Should().BeEmpty();
        }

        [Fact]
        public void Width_returns_shape_width_in_pixels()
        {
            // Arrange
            IShape shapeCase1 = SCPresentation.Open(GetTestStream("006_1 slides.pptx")).Slides[0].Shapes.First(sp => sp.Id == 2);
            IGroupShape groupShape = (IGroupShape)SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 7);
            IShape shapeCase2 = groupShape.Shapes.First(sp => sp.Id == 5);
            IShape shapeCase3 = SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 9);

            // Act
            int width1 = shapeCase1.Width;
            int width2 = shapeCase2.Width;
            int width3 = shapeCase3.Width;

            // Assert
            (width1 * 914400 / Helpers.TestHelper.HorizontalResolution).Should().Be(9144000);
            (width2 * 914400 / Helpers.TestHelper.HorizontalResolution).Should().Be(1181100);
            (width3 * 914400 / Helpers.TestHelper.HorizontalResolution).Should().Be(485775);
        }

        [Theory]
        [InlineData("050_title-placeholder.pptx", 1, 2, 777)]
        [InlineData("051_title-placeholder.pptx", 1, 3074, 864)]
        public void Width_returns_width_of_Title_placeholder(string filename, int slideNumber, int shapeId,
            int expectedWidth)
        {
            // Arrange
            var pptx = GetTestStream(filename);
            var pres = SCPresentation.Open(pptx);
            var autoShape = pres.Slides[slideNumber - 1].Shapes.GetById<IAutoShape>(shapeId);

            // Act
            var shapeWidth = autoShape.Width;

            // Assert
            shapeWidth.Should().Be(expectedWidth);
        }

        [Fact]
        public void Height_ReturnsHeightInPixels()
        {
            // Arrange
            IShape shapeCase1 = SCPresentation.Open(GetTestStream("006_1 slides.pptx")).Slides[0].Shapes.First(sp => sp.Id == 2);
            IGroupShape groupShape = (IGroupShape)SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 7);
            IShape shapeCase2 = groupShape.Shapes.First(sp => sp.Id == 5);
            IShape shapeCase3 = SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 9);
            float verticalResulution = Helpers.TestHelper.VerticalResolution;

            // Act
            int height1 = shapeCase1.Height;
            int height2 = shapeCase2.Height;
            int height3 = shapeCase3.Height;

            // Assert
            (height1 * 914400 / verticalResulution).Should().Be(1419225);
            (height2 * 914400 / verticalResulution).Should().Be(647700);
            (height3 * 914400 / verticalResulution).Should().Be(371475);
        }

        [Theory]
        [MemberData(nameof(GeometryTypeTestCases))]
        public void GeometryType_returns_shape_geometry_type(IShape shape, SCGeometry expectedGeometryType)
        {
            // Assert
            shape.GeometryType.Should().Be(expectedGeometryType);
        }

        public static IEnumerable<object[]> GeometryTypeTestCases()
        {
            var pptxStream = GetTestStream("021.pptx");
            var presentation = SCPresentation.Open(pptxStream);
            var shapeCase1 = presentation.Slides[3].Shapes.First(sp => sp.Id == 2);
            var shapeCase2 = presentation.Slides[3].Shapes.First(sp => sp.Id == 3);

            yield return new object[] { shapeCase1, SCGeometry.Rectangle };
            yield return new object[] { shapeCase2, SCGeometry.Ellipse };
        }

        [Fact]
        public void Shape_IsOLEObject()
        {
            // Arrange
            IOLEObject oleObject = SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 8) as IOLEObject;

            // Act-Assert
            oleObject.Should().NotBeNull();
        }

        [Fact]
        public void Shape_IsNotGroupShape()
        {
            // Arrange
            IShape shape = SCPresentation.Open(GetTestStream("006_1 slides.pptx")).Slides[0].Shapes.First(x => x.Id == 3);

            // Act-Assert
            shape.Should().NotBeOfType<IGroupShape>();
        }

        [Fact]
        public void Shape_IsNotAutoShape()
        {
            // Arrange
            var pptx9 = GetTestStream("009_table.pptx");
            var pres9 = SCPresentation.Open(pptx9);
            var pptx11 = GetTestStream("011_dt.pptx");
            var pres11 = SCPresentation.Open(pptx11);
            var shapeCase1 = pres9.Slides[4].Shapes.First(sp => sp.Id == 5);
            var shapeCase2 = pres11.Slides[0].Shapes.First(sp => sp.Id == 4);

            // Act-Assert
            shapeCase1.Should().NotBeOfType<IAutoShape>();
            shapeCase2.Should().NotBeOfType<IAutoShape>();
        }

        [Fact]
        public void CustomData_ReturnsNull_WhenShapeHasNotCustomData()
        {
            // Arrange
            var shape = SCPresentation.Open(GetTestStream("009_table.pptx")).Slides.First().Shapes.First();

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
            var presentation = SCPresentation.Open(GetTestStream("009_table.pptx"));
            var shape = presentation.Slides.First().Shapes.First();

            // Act
            shape.CustomData = customDataString;
            presentation.SaveAs(savedPreStream);

            // Assert
            presentation = SCPresentation.Open(savedPreStream);
            shape = presentation.Slides.First().Shapes.First();
            shape.CustomData.Should().Be(customDataString);
        }

        [Fact]
        public void Name_ReturnsShapeNameString()
        {
            // Arrange
            IShape shape = SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 8);

            // Act
            string shapeName = shape.Name;

            // Assert
            shapeName.Should().BeEquivalentTo("Object 2");
        }

        [Fact]
        public void Hidden_ReturnsValueIndicatingWhetherShapeIsHiddenFromTheSlide()
        {
            // Arrange
            IShape shapeCase1 = SCPresentation.Open(GetTestStream("004.pptx")).Slides[0].Shapes[0];
            IShape shapeCase2 = SCPresentation.Open(GetTestStream("004.pptx")).Slides[0].Shapes[1];

            // Act-Assert
            shapeCase1.Hidden.Should().BeTrue();
            shapeCase2.Hidden.Should().BeFalse();
        }
    }
}

#endif