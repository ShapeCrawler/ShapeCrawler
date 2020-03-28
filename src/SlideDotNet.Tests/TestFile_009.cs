using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using NSubstitute;
using SlideDotNet.Enums;
using SlideDotNet.Extensions;
using SlideDotNet.Models;
using SlideDotNet.Models.Settings;
using SlideDotNet.Models.SlideComponents;
using SlideDotNet.Services;
using SlideDotNet.Services.Placeholders;
using SlideDotNet.Tests.Helpers;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;
using System.Linq;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace SlideDotNet.Tests
{
    public class TestFile_009
    {
        [Fact]
        public void SlidesNumber_Test()
        {
            var ms = new MemoryStream(Properties.Resources._001);
            var pre = new PresentationEx(ms);

            // ACT
            var sldNumber = pre.Slides.Count();

            // CLOSE
            pre.Close();
            
            // ASSERT
            Assert.Equal(2, sldNumber);
        }

        /// <State>
        /// - there is a presentation with two slides. The first slide contains one element. The second slide includes two elements;
        /// - the first slide is removed;
        /// - the presentation is closed.
        /// </State>
        /// <ExpectedBahavior>
        /// There is the only second slide with its two elements.
        /// </ExpectedBahavior>
        [Fact]
        public void SlidesRemove_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._007_2_slides);
            var pre = new PresentationEx(ms);

            // ACT
            var slide1 = pre.Slides.First();
            pre.Slides.Remove(slide1);
            pre.Close();

            var pre2 = new PresentationEx(ms);
            var numSlides = pre2.Slides.Count();
            var numElements = pre2.Slides.Single().Shapes.Count;
            pre2.Close();
            ms.Dispose();

            // ASSERT
            Assert.Equal(1, numSlides);
            Assert.Equal(2, numElements);
        }

        [Fact]
        public void ShapeTextBody_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._008);

            // ACT
            var shapes = pre.Slides.Single().Shapes.OfType<ShapeEx>();
            var sh36 = shapes.Single(e => e.Id == 36);
            var sh37 = shapes.Single(e => e.Id == 37);
           
            pre.Close();

            // ASSERT
            Assert.False(sh36.HasTextFrame);
            Assert.True(sh37.HasTextFrame);
            Assert.Equal("P1t1 P1t2", sh37.TextFrame.Paragraphs[0].Text);
            Assert.Equal("p2", sh37.TextFrame.Paragraphs[1].Text);
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
        public void GroupsElementPropertiesTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);

            // ACT
            var groupElement = pre.Slides[1].Shapes.Single(x => x.ContentType.Equals(ShapeContentType.Group));
            var groupedShape5 = groupElement.GroupedShapes.Single(x => x.Id.Equals(5));
            pre.Close();

            // ASSERT
            Assert.Equal(1581846, groupedShape5.X);
            Assert.Equal(1181377, groupedShape5.Width);
            Assert.Equal(654096, groupedShape5.Height);
        }

        [Fact]
        public void SecondSlideElementsNumberTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);

            // ACT
            var elNumber1 = pre.Slides[0].Shapes.Count;
            var elNumber2 = pre.Slides[1].Shapes.Count;
            pre.Close();

            // ASSERT
            Assert.Equal(6, elNumber1);
            Assert.Equal(6, elNumber2);
        }

        [Fact]
        public void SlideElementsDoNotThrowsExceptionTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);

            // ACT
            var elements = pre.Slides[0].Shapes;

            pre.Close();
        }

        [Fact]
        public void PictureEx_BytesTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
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
            pre.Close();
            testImage2Stream.Dispose();

            // ASSERT
            Assert.NotEqual(sizeBefore,  sizeAfter);
        }

        [Fact]
        public void Shape_Fill_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
            var sp4 = pre.Slides[2].Shapes.Single(e => e.Id.Equals(4));

            // ACT
            var fillType = sp4.Fill.Type;
            var fillPicLength = sp4.Fill.Picture.GetImageBytes().Result.Length;

            // ASSERT
            Assert.Equal(FillType.Picture, fillType);
            Assert.True(fillPicLength > 0);
        }

        [Fact]
        public void Shape_Fill_Solid_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
            var sp2 = pre.Slides[1].Shapes.Single(e => e.Id.Equals(2));

            // ACT
            var fillType = sp2.Fill.Type;
            var fillSolidColorName = sp2.Fill.SolidColor.Name;

            // ASSERT
            Assert.Equal(FillType.Solid, fillType);
            Assert.Equal("ff0000", fillSolidColorName);
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
        public void Shape_Fill_IsNull_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
            var shapeEx = pre.Slides[1].Shapes.Single(e => e.Id.Equals(6));

            // ACT
            var shapeFill = shapeEx.Fill;

            // ASSERT
            Assert.Null(shapeFill);
        }

        [Fact]
        public void OleObjects_ParseTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
            var shapes = pre.Slides[1].Shapes;

            // ACT
            var oleNumbers = shapes.Count(e => e.ContentType.Equals(ShapeContentType.OLEObject));
            var ole9 = shapes.Single(s => s.Id == 9);

            pre.Close();

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
            var pre = new PresentationEx(Properties.Resources._009);

            // ACT
            var name = pre.Slides[1].Shapes.Single(e => e.Id.Equals(8)).Name;

            pre.Close();

            // ASSERT
            Assert.Equal("Object 2", name);
        }

        [Fact]
        public void SlideEx_Background_IsNullTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);

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
            var pre = new PresentationEx(Properties.Resources._009);
            var shape = (ShapeEx)pre.Slides[2].Shapes.SingleOrDefault(e => e.Id.Equals(2));
            var paragraphs = shape.TextFrame.Paragraphs;

            // ACT
            var numParagraphs = paragraphs.Count;
            var portions = paragraphs[0].Portions;
            var numPortions = portions.Count;
            var por1Size = portions[0].FontHeight;
            var por2Size = portions[1].FontHeight;

            pre.Close();

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
            var pre = new PresentationEx(Properties.Resources._009);

            // ACT
            var w = pre.SlideWidth;
            var y = pre.SlideHeight;

            pre.Close();

            // ASSERT
            Assert.Equal(9144000, w);
            Assert.Equal(5143500, y);
        }

        [Fact]
        public void Placeholder_FontHeight_TextBox_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
            var elements = pre.Slides[3].Shapes;
            var tb2TitlePh = elements.Single(e => e.Id.Equals(2));
            var subTitle3 = elements.Single(e => e.Id.Equals(3));

            // ACT
            var fhTitle = tb2TitlePh.TextFrame.Paragraphs.Single().Portions.Single().FontHeight;
            var text2 = tb2TitlePh.TextFrame.Text;
            var fhSubTitle = subTitle3.TextFrame.Paragraphs.Single().Portions.Single().FontHeight;

            pre.Close();

            // ASSERT
            Assert.Equal(4400, fhTitle);
            Assert.Equal(3200, fhSubTitle);
            Assert.Equal("Title text", text2);
        }

        [Fact]
        public void TablesPropertiesTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
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

            pre.Close();

            // ASSERT
            Assert.Equal(3, numRows);
            Assert.Equal(3, numCells);
            Assert.Equal(2, numParagraphs);
            Assert.Equal("0:0_p1_lvl1\r\n0:0_p2_lvl2", cellTxt);
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
            var pre = new PresentationEx(Properties.Resources._009);
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

            pre.Close();

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
            var pre = new PresentationEx(Properties.Resources._009);
            var sld3Shapes = pre.Slides[2].Shapes;
            var pieChartShape7 = sld3Shapes.Single(e => e.Id == 7);
            var pieChart7 = pieChartShape7.Chart;
            var pieChart7Categories = pieChart7.Categories;

            // ACT
            var c1 = pieChart7Categories[0];
            var c2 = pieChart7Categories[1];
            var c3 = pieChart7Categories[2];
            var c4 = pieChart7Categories[3];

            pre.Close();

            // ASSERT
            Assert.Equal("Q1", c1);
            Assert.Equal("Q2", c2);
            Assert.Equal("Q3", c3);
            Assert.Equal("Q4", c4);
        }

        [Fact]
        public void DateTimePlaceholder_HasTextFrame_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._008);
            var sp3 = pre.Slides[0].Shapes.Single(sp => sp.Id == 3);

            // ACT
            var hasTextFrame = sp3.HasTextFrame;
            var text = sp3.TextFrame.Text;
            var phType = sp3.PlaceholderType;

            pre.Close();

            // ASSERT
            Assert.True(hasTextFrame);
            Assert.Equal("25.01.2020", text);
            Assert.Equal(PlaceholderType.DateAndTime, phType);
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

        [Fact]
        public void ElementsNumber()
        {
            var ms = new MemoryStream(Properties.Resources._002);
            var pre = new PresentationEx(ms);
            var allElements = pre.Slides.First().Shapes;

            // ACT
            var elementsNumber = allElements.Count;

            // CLOSE
            pre.Close();

            // ASSERT
            Assert.Equal(3, elementsNumber);
        }

        /// <State>
        /// - there is presentation with two slides;
        /// - first slide is deleted.
        /// </State>
        /// <ExpectedBahavior>
        /// Presentation contains single slide with 1 number.
        /// </ExpectedBahavior>
        [Fact]
        public void Remove_Test1()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._007_2_slides);
            var slides = pre.Slides;
            var slide1 = slides[0];
            var slide2 = slides[1];

            // ACT
            var num1BeforeRemoving = slide1.Number;
            var num2BeforeRemoving = slide2.Number;
            slides.Remove(slide1);
            var num2AfterRemoving = slide2.Number;

            // ARRANGE
            Assert.Equal(1, num1BeforeRemoving);
            Assert.Equal(2, num2BeforeRemoving);
            Assert.Equal(1, num2AfterRemoving);

            // CLEAN
            pre.Close();
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
        public void Get_Test()
        {
            var ms = new MemoryStream(Properties.Resources._008);
            var xmlDoc = PresentationDocument.Open(ms, false);
            var sldPart = xmlDoc.PresentationPart.SlideParts.First();
            var spId3 = sldPart.Slide.CommonSlideData.ShapeTree.Elements<DocumentFormat.OpenXml.Presentation.Shape>().Single(sp => sp.GetId() == 3);
            var sldLtPart = sldPart.SlideLayoutPart;
            var phService = new PlaceholderService(sldLtPart);

            // ACT
            var type = phService.TryGet(spId3).PlaceholderType;

            // CLOSE
            xmlDoc.Close();

            // ASSERT
            Assert.Equal(PlaceholderType.DateAndTime, type);
        }

        [Fact]
        public void GetPlaceholderType_Test()
        {
            var ms = new MemoryStream(Properties.Resources._008);
            var xmlDoc = PresentationDocument.Open(ms, false);
            var sldPart = xmlDoc.PresentationPart.SlideParts.First();
            var spId3 = sldPart.Slide.CommonSlideData.ShapeTree.Elements<DocumentFormat.OpenXml.Presentation.Shape>().Single(sp => sp.GetId() == 3);

            // ACT
            var phXml = PlaceholderService.PlaceholderDataFrom(spId3);

            // CLOSE
            xmlDoc.Close();

            // ASSERT
            Assert.Equal(PlaceholderType.DateAndTime, phXml.PlaceholderType);
        }

        [Fact]
        public void CreateShapesCollection_Test()
        {
            // ARRANGE
            var ms = new MemoryStream(Properties.Resources._003);
            var doc = PresentationDocument.Open(ms, false);

            var xmlSldPart = doc.PresentationPart.SlideParts.First();
            var preSettings = new PreSettings(doc.PresentationPart.Presentation);
            var shapeTree = xmlSldPart.Slide.CommonSlideData.ShapeTree;
            var parser = new ShapeFactory(xmlSldPart, preSettings);

            // ACT
            var candidates = parser.FromTree(shapeTree);

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
