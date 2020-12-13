using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using SlideDotNet.Enums;
using SlideDotNet.Extensions;
using SlideDotNet.Models;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using Xunit;
using System.Linq;
using FluentAssertions;
using SlideDotNet.Services.ShapeCreators;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace SlideDotNet.Tests
{
    public class TestFile_009Fixture : IDisposable
    {
        public PresentationEx pre009 { get; }

        public TestFile_009Fixture()
        {
            pre009 = new PresentationEx(Properties.Resources._009);
        }

        public void Dispose()
        {
            pre009.Close();
        }
    }

    public class TestFile_009 : IClassFixture<TestFile_009Fixture>
    {
        private readonly TestFile_009Fixture _fixture;

        public TestFile_009(TestFile_009Fixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void SlideElementsCount_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._003);

            // ACT
            var numberElements = pre.Slides.Single().Shapes.Count;
            pre.Close();

            // ASSERT
            Assert.Equal(5, numberElements);
        }

        [Fact]
        public void TextBox_Placeholder_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._006_1_slides);

            // ACT
            var shapePlaceholder = pre.Slides[0].Shapes.First(x => x.Id == 2);
            pre.Close();

            // ASSERT
            Assert.Equal(1524000, shapePlaceholder.X);
            Assert.Equal(1122363, shapePlaceholder.Y);
            Assert.Equal(9144000, shapePlaceholder.Width);
            Assert.Equal(1425528, shapePlaceholder.Height);
        }

        [Fact]
        public void Shape_XandWsetter_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._006_1_slides);
            var shape2 = pre.Slides[0].Shapes.First(x => x.Id == 3);

            // ACT
            shape2.X = 4000000;
            shape2.Width = 6000000;
            var isPlaceholder = shape2.IsPlaceholder;
            var isGrouped = shape2.IsGrouped;

            var ms = new MemoryStream();
            pre.SaveAs(ms);
            pre.Close();

            ms.SeekBegin();
            pre = new PresentationEx(ms);
            shape2 = pre.Slides[0].Shapes.First(x => x.Id == 3);
            pre.Close();

            // ASSERT
            Assert.False(isPlaceholder);
            Assert.False(isGrouped);
            Assert.Equal(4000000, shape2.X);
            Assert.Equal(6000000, shape2.Width);
        }

        [Fact]
        public void GroupedShape()
        {
            // Arrange
            var pre = _fixture.pre009;
            var groupShape = pre.Slides[1].Shapes.Single(x => x.ContentType.Equals(ShapeContentType.Group));

            // Act
            var groupedShape = groupShape.GroupedShapes.Single(x => x.Id.Equals(5));

            // Assert
            Assert.Equal(1581846, groupedShape.X);
            Assert.Equal(1181377, groupedShape.Width);
            Assert.Equal(654096, groupedShape.Height);
        }

        [Fact]
        public void ShapeCustomData_ShouldReturnNull_ShapeCustomDataIsNotSet()
        {
            // Arrange
            var pre = _fixture.pre009;
            var shape = pre.Slides.First().Shapes.First();

            // Act
            var shapeCustomData = shape.CustomData;

            // Assert
            shapeCustomData.Should().BeNull();
        }

        [Fact]
        public void ShapeCustomData_ShouldReturnData_ShapeCustomDataIsSet()
        {
            // Arrange
            const string customDataString = "Test custom data";
            var origPreStream = new MemoryStream();
            origPreStream.Write(Properties.Resources._009);
            var originVersionPre = new PresentationEx(origPreStream);
            var shape = originVersionPre.Slides.First().Shapes.First();

            // Act
            shape.CustomData = customDataString;
            var changedVersionPreStream = new MemoryStream();
            originVersionPre.SaveAs(changedVersionPreStream);
            var changedVersionPre = new PresentationEx(changedVersionPreStream);
            var shapeCustomData = changedVersionPre.Slides.First().Shapes.First().CustomData;

            // Assert
            shapeCustomData.Should().Be(customDataString);
        }

        [Fact]
        public void SecondSlideElementsNumberTest()
        {
            // ARRANGE
            var pre = _fixture.pre009;

            // ACT
            var elNumber1 = pre.Slides[0].Shapes.Count;
            var elNumber2 = pre.Slides[1].Shapes.Count;

            // ASSERT
            Assert.Equal(6, elNumber1);
            Assert.Equal(6, elNumber2);
        }

        [Fact]
        public void SlideElementsDoNotThrowsExceptionTest()
        {
            // ARRANGE
            var pre = _fixture.pre009;

            // ACT
            var elements = pre.Slides[0].Shapes;
        }

        [Fact]
        public void PictureEx_BytesTest()
        {
            // ARRANGE
            var pre = _fixture.pre009;
            var picEx = pre.Slides[1].Shapes.Single(e => e.Id.Equals(3));

            // ACT
            var hasPicture = picEx.HasPicture;
            var bytes = picEx.Picture.ImageEx.GetImageBytes().Result;

            // ASSERT
            Assert.True(bytes.Length > 0);
            Assert.True(hasPicture);
        }

        [Fact]
        public void PictureEx_SetImageTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
            var picEx = pre.Slides[1].Shapes.Single(e => e.Id.Equals(3));
            var testImage2Stream = new MemoryStream(Properties.Resources.test_image_2);
            var sizeBefore = picEx.Picture.ImageEx.GetImageBytes().Result.Length;

            // ACT
            picEx.Picture.ImageEx.SetImageStream(testImage2Stream);

            var sizeAfter = picEx.Picture.ImageEx.GetImageBytes().Result.Length;

            // ASSERT
            Assert.NotEqual(sizeBefore,  sizeAfter);
        }

        [Fact]
        public void Shape_Fill_Test()
        {
            // ARRANGE
            var pre = _fixture.pre009;
            var sp4 = pre.Slides[2].Shapes.Single(e => e.Id.Equals(4));

            // ACT
            var fillType = sp4.Fill.Type;
            var fillPicLength = sp4.Fill.Picture.GetImageBytes().Result.Length;

            // ASSERT
            Assert.Equal(FillType.Picture, fillType);
            Assert.True(fillPicLength > 0);
        }

        [Fact]
        public void ShapeFill_FillTypeAndFillSolidColorName()
        {
            // Arrange
            var pre = _fixture.pre009;
            var sp2 = pre.Slides[1].Shapes.Single(e => e.Id.Equals(2));
            var shapeFill = sp2.Fill;

            // Act
            var fillType = shapeFill.Type;
            var fillSolidColorName = shapeFill.SolidColor.Name;

            // Assert
            fillType.Should().BeEquivalentTo(FillType.Solid);
            fillSolidColorName.Should().BeEquivalentTo("ff0000");
        }


        [Fact]
        public void ShapeFill_ShouldReturnNull_ShapeIsNotFilled()
        {
            // Arrange
            var pre = _fixture.pre009;
            var shapeEx = pre.Slides[1].Shapes.Single(e => e.Id.Equals(6));

            // Act
            var shapeFill = shapeEx.Fill;

            // Act
            shapeFill.Should().BeNull();
        }

        [Fact]
        public void ShapeEx_BackgroundImage_SetImageTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
            var shapeEx = (ShapeEx)pre.Slides[2].Shapes.Single(e => e.Id.Equals(4));
            var testImage2Stream = new MemoryStream(Properties.Resources.test_image_2);
            var sizeBefore = shapeEx.Fill.Picture.GetImageBytes().Result.Length;

            // ACT
            shapeEx.Fill.Picture.SetImageStream(testImage2Stream);

            var sizeAfter = shapeEx.Fill.Picture.GetImageBytes().Result.Length;
            pre.Close();
            testImage2Stream.Dispose();

            // ASSERT
            Assert.NotEqual(sizeBefore, sizeAfter);
        }

        [Fact]
        public void OleObjects_ParseTest()
        {
            // ARRANGE
            var pre = _fixture.pre009;
            var shapes = pre.Slides[1].Shapes;

            // ACT
            var oleNumbers = shapes.Count(e => e.ContentType.Equals(ShapeContentType.OLEObject));
            var ole9 = shapes.Single(s => s.Id == 9);

            // ASSERT
            Assert.Equal(2, oleNumbers);
            Assert.Equal(699323, ole9.X);
            Assert.Equal(3463288, ole9.Y);
            Assert.Equal(485775, ole9.Width);
            Assert.Equal(373062, ole9.Height);
        }

        [Fact]
        public void OLEObject_NameTest()
        {
            // ARRANGE
            var pre = _fixture.pre009;

            // ACT
            var name = pre.Slides[1].Shapes.Single(e => e.Id.Equals(8)).Name;

            // ASSERT
            Assert.Equal("Object 2", name);
        }

        [Fact]
        public void SlideEx_Background_IsNullTest()
        {
            // ARRANGE
            var pre = _fixture.pre009;

            // ACT
            var bg = pre.Slides[1].BackgroundImage;

            // ASSERT
            Assert.Null(bg);
        }

        [Fact]
        public void SlideEx_Background_ChangeTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
            var bg = pre.Slides[0].BackgroundImage;
            var testImage2Stream = new MemoryStream(Properties.Resources.test_image_2);
            var sizeBefore = bg.GetImageBytes().Result.Length;

            // ACT
            bg.SetImageStream(testImage2Stream);

            var sizeAfter = bg.GetImageBytes().Result.Length;
            pre.Close();
            testImage2Stream.Dispose();

            // ASSERT
            Assert.NotEqual(sizeBefore, sizeAfter);
        }

        [Fact]
        public void NumberParagraphAndPortionTest()
        {
            // ARRANGE
            var pre = _fixture.pre009;
            var shape = (ShapeEx)pre.Slides[2].Shapes.SingleOrDefault(e => e.Id.Equals(2));
            var paragraphs = shape.TextFrame.Paragraphs;

            // ACT
            var numParagraphs = paragraphs.Count;
            var portions = paragraphs[0].Portions;
            var numPortions = portions.Count;
            var por1Size = portions[0].FontHeight;
            var por2Size = portions[1].FontHeight;


            // ASSERT
            Assert.Equal(1, numParagraphs);
            Assert.Equal(2, numPortions);
            Assert.Equal(1800, por1Size);
            Assert.Equal(2000, por2Size);
        }

        [Fact]
        public void SlideWidthAndHeightTest()
        {
            // ARRANGE
            var pre = _fixture.pre009;

            // ACT
            var w = pre.SlideWidth;
            var y = pre.SlideHeight;

            // ASSERT
            Assert.Equal(9144000, w);
            Assert.Equal(5143500, y);
        }

        [Fact]
        public void Placeholder_FontHeight_TextBox_Test()
        {
            // ARRANGE
            var pre = _fixture.pre009;
            var elements = pre.Slides[3].Shapes;
            var tb2TitlePh = elements.Single(e => e.Id.Equals(2));
            var subTitle3 = elements.Single(e => e.Id.Equals(3));

            // ACT
            var fhTitle = tb2TitlePh.TextFrame.Paragraphs.Single().Portions.Single().FontHeight;
            var text2 = tb2TitlePh.TextFrame.Text;
            var fhSubTitle = subTitle3.TextFrame.Paragraphs.Single().Portions.Single().FontHeight;

            // ASSERT
            Assert.Equal(4400, fhTitle);
            Assert.Equal(3200, fhSubTitle);
            Assert.Equal("Title text", text2);
        }

        [Fact]
        public void TablesPropertiesTest()
        {
            // ARRANGE
            var pre = _fixture.pre009;
            var elements = pre.Slides[2].Shapes;
            var tblEx = elements.Single(e => e.Id.Equals(3));
            var firstRow = tblEx.Table.Rows.First();

            // ACT
            var numRows = tblEx.Table.Rows.Count;
            var numCells = firstRow.Cells.Count;
            var numParagraphs = firstRow.Cells.First().TextBody.Paragraphs.Count;
            var cellTxt = firstRow.Cells.First().TextBody.Text;
            var prText = firstRow.Cells.First().TextBody.Paragraphs.First().Text;
            var portionTxt = firstRow.Cells.First().TextBody.Paragraphs.First().Portions.Single().Text;

            // ASSERT
            Assert.Equal(3, numRows);
            Assert.Equal(3, numCells);
            Assert.Equal(2, numParagraphs);
            Assert.Equal("0:0_p1_lvl1", prText);
            Assert.Equal("0:0_p1_lvl1", portionTxt);
        }

        [Fact]
        public void Table_Row_Remove_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
            var sld3Shapes = pre.Slides[2].Shapes;
            var table3 = sld3Shapes.First(s => s.Id.Equals(3)).Table;
            var rows = table3.Rows;
            var numRowsBefore = rows.Count;

            // ACT
            rows.RemoveAt(0);
            
            var ms = new MemoryStream();
            pre.SaveAs(ms);
            pre.Close();

            pre = new PresentationEx(ms);
            table3 = pre.Slides[2].Shapes.First(s => s.Id.Equals(3)).Table;
            rows = table3.Rows;
            var numRowsAfter = rows.Count;

            // ASSERT
            Assert.True(numRowsBefore > numRowsAfter);
        }

        [Fact]
        public void ChartPropertiesTest()
        {
            // ARRANGE
            var pre = _fixture.pre009;
            var sld3Elements = pre.Slides[2].Shapes;
            var sld5Elements = pre.Slides[4].Shapes;
            var chartEx6 = sld3Elements.Single(e => e.Id.Equals(6));
            var pieChartShape7 = sld3Elements.Single(e => e.Id.Equals(7));
            var sld5Chart6 = sld5Elements.Single(e => e.Id.Equals(6));
            var sld5Chart3 = sld5Elements.Single(e => e.Id.Equals(3));
            var sld5Chart5 = sld5Elements.Single(e => e.Id.Equals(5));
            var pieChart7 = pieChartShape7.Chart;

            // ACT
            var pieChart7Title = pieChart7.Title;
            var pieChart7Type = pieChart7.Type;
            var pieChart7NumSeries = pieChart7.SeriesCollection.Count();
            var chart6Title = chartEx6.Chart.Title;
            var sld5Chart6Title = sld5Chart6.Chart.Title;
            var sld5Chart3Title = sld5Chart3.Chart.Title;
            var sld5Chart5Title = sld5Chart5.Chart.Title;
            var hasTextFrame = sld5Chart5.HasTextFrame;
            var v00 = pieChart7.SeriesCollection[0].PointValues[0];
            var v01 = pieChart7.SeriesCollection[0].PointValues[1];

            // ASSERT
            Assert.Equal("Sales", pieChart7Title);
            Assert.Equal("Sales2", chart6Title);
            Assert.Equal("Sales3", sld5Chart6Title);
            Assert.Equal("Sales4", sld5Chart3Title);
            Assert.Equal("Sales5", sld5Chart5Title);
            Assert.Equal(ChartType.PieChart, pieChart7Type);
            Assert.False(hasTextFrame);
            Assert.Equal(1, pieChart7NumSeries);
            Assert.Equal(8.2, v00);
            Assert.Equal(3.2, v01);
        }

        [Fact]
        public void Chart_Category_Test()
        {
            // ARRANGE
            var pre = _fixture.pre009;
            var sld3Shapes = pre.Slides[2].Shapes;
            var pieChartShape7 = sld3Shapes.Single(e => e.Id == 7);
            var pieChart7 = pieChartShape7.Chart;
            var pieChart7Categories = pieChart7.Categories;

            // ACT
            var c1 = pieChart7Categories[0].Name;
            var c2 = pieChart7Categories[1].Name;
            var c3 = pieChart7Categories[2].Name;
            var c4 = pieChart7Categories[3].Name;

            // ASSERT
            Assert.Equal("Q1", c1);
            Assert.Equal("Q2", c2);
            Assert.Equal("Q3", c3);
            Assert.Equal("Q4", c4);
        }

        [Fact]
        public void DateTimePlaceholder_Text_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._011_dt);
            var dt = pre.Slides[0].Shapes.Single(s => s.Id == 54275);

            // ACT
            var text = dt.TextFrame.Text;
            var hasText = dt.HasTextFrame;

            pre.Close();

            // ASSERT
            Assert.True(hasText);
            Assert.NotNull(text);
        }

        [Fact]
        public void Placeholder_FontHeight_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._012_title_placeholder);
            var title = pre.Slides[0].Shapes.Single(x => x.Id == 2);

            // ACT
            var text = title.TextFrame.Text;
            var fh = title.TextFrame.Paragraphs.First().Portions.First().FontHeight;

            pre.Close();

            // ASSERT
            Assert.NotNull(text);
            Assert.Equal(2000, fh);
        }

        [Fact]
        public void Placeholder_FontHeight_Title_Test()
        {
            // ARRANGE
            var pre010 = new PresentationEx(Properties.Resources._010);
            var pre010TextBox = pre010.Slides[0].Shapes.Single(x => x.Id == 2);

            // ACT
            var fh = pre010TextBox.TextFrame.Paragraphs.First().Portions.First().FontHeight;

            pre010.Close();

            // ASSERT
            Assert.Equal(1539, fh);
        }

        /// <State>
        /// - there is a single slide presentation;
        /// - slide is deleted.
        /// </State>
        /// <ExpectedBahavior>
        /// Presentation is empty.
        /// </ExpectedBahavior>
        [Fact]
        public void Remove_Test2()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._006_1_slides);
            var slides = pre.Slides;
            var slide1 = slides.First();

            // ACT
            slides.Remove(slide1);

            // ARRANGE
            Assert.Empty(slides);

            // CLEAN
            pre.Close();
        }

        [Fact]
        public void Shape_Text_Test2()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._012_title_placeholder);
            var autoShape = pre.Slides[0].Shapes.Single(s => s.Id == 3);

            // ACT
            var text = autoShape.TextFrame.Text;

            pre.Close();

            // ASSERT
            Assert.Equal("P1 P2", text);
        }

        [Fact]
        public void Shape_Text_Tests()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._011_dt);
            var autoShape = pre.Slides[0].Shapes.Single(s => s.Id == 2);
            var grShape = pre.Slides[0].Shapes.Single(s => s.Id == 4);

            // ACT
            var text = autoShape.TextFrame.Text;
            var hasTextFrame = grShape.HasTextFrame;

            pre.Close();

            // ASSERT
            Assert.NotNull(text);
            Assert.False(hasTextFrame);
        }

        [Fact]
        public void Hidden_Test()
        {
            var ms = new MemoryStream(Properties.Resources._004);
            var pre = new PresentationEx(ms);

            // ACT
            var allElements = pre.Slides.Single().Shapes;
            var shapeHiddenValue = allElements[0].Hidden;
            var tableHiddenValue = allElements[1].Hidden;

            // CLOSE
            pre.Close();

            // ASSERT
            Assert.True(shapeHiddenValue);
            Assert.False(tableHiddenValue);
        }
        [Fact]
        public void CreateShapesCollection_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._003);
            var doc = PresentationDocument.Open(ms, false);

            var sdkSldPart = doc.PresentationPart.SlideParts.First();
            var preSettings = new PreSettings(doc.PresentationPart.Presentation, new Lazy<SlideSize>());
            var parser = new ShapeFactory(preSettings);

            // ACT
            var candidates = parser.FromSldPart(sdkSldPart);

            // CLEAN
            doc.Dispose();
            ms.Dispose();

            // ASSERT
            Assert.Single(candidates.Where(c => c.ContentType.Equals(ShapeContentType.AutoShape)));
            Assert.Single(candidates.Where(c => c.ContentType.Equals(ShapeContentType.Picture)));
            Assert.Single(candidates.Where(c => c.ContentType.Equals(ShapeContentType.Table)));
            Assert.Single(candidates.Where(c => c.ContentType.Equals(ShapeContentType.Chart)));
            Assert.Single(candidates.Where(c => c.ContentType.Equals(ShapeContentType.Group)));
        }
    }
}
