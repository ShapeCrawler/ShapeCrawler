using System.IO;
using System.Linq;
using SlideXML.Enums;
using SlideXML.Models;
using SlideXML.Models.SlideComponents;
using Xunit;

namespace SlideXML.Tests
{
    /// <summary>
    /// Represents tests of the <see cref="PresentationSL"/> class.
    /// </summary>
    public class PresentationExTests
    {
        [Fact]
        public void SlidesNumber_Test()
        {
            var ms = new MemoryStream(Properties.Resources._001);
            var pre = new PresentationSL(ms);

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
            var pre = new PresentationSL(ms);

            // ACT
            var slide1 = pre.Slides.First();
            pre.Slides.Remove(slide1);
            pre.Close();

            var pre2 = new PresentationSL(ms);
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
            var pre = new PresentationSL(Properties.Resources._008);

            // ACT
            var shapes = pre.Slides.Single().Shapes.OfType<ShapeSL>();
            var sh36 = shapes.Single(e => e.Id == 36);
            var sh37 = shapes.Single(e => e.Id == 37);
            pre.Close();

            // ASSERT
            Assert.Null(sh36.TextFrame);
            Assert.NotNull(sh37.TextFrame);
            Assert.Equal("P1t1 P1t2", sh37.TextFrame.Paragraphs[0].Text);
            Assert.Equal("p2", sh37.TextFrame.Paragraphs[1].Text);
        }

        [Fact]
        public void SlideElementsCount_Test()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._003);

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
            var pre = new PresentationSL(Properties.Resources._006_1_slides);

            // ACT
            var shapePlaceholder = pre.Slides.Single().Shapes.Single();
            pre.Close();

            // ASSERT
            Assert.Equal(1524000, shapePlaceholder.X);
            Assert.Equal(1122363, shapePlaceholder.Y);
            Assert.Equal(9144000, shapePlaceholder.Width);
            Assert.Equal(2387600, shapePlaceholder.Height);
        }

        [Fact]
        public void GroupsElementPropertiesTest()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._009);

            // ACT
            var slides = pre.Slides;
            var groupElement = pre.Slides[1].Shapes.Single(x => x.Type.Equals(ShapeType.Group));
            var el3 = groupElement.Group.Shapes.Single(x => x.Id.Equals(5));
            pre.Close();

            // ASSERT
            Assert.Equal(1581846, el3.X);
            Assert.Equal(1181377, el3.Width);
            Assert.Equal(654096, el3.Height);
        }

        [Fact]
        public void SecondSlideElementsNumberTest()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._009);

            // ACT
            var elNumber1 = pre.Slides[0].Shapes.Count;
            var elNumber2 = pre.Slides[1].Shapes.Count;
            pre.Close();

            // ASSERT
            Assert.Equal(6, elNumber1);
            Assert.Equal(5, elNumber2);
        }

        [Fact]
        public void SlideElementsDoNotThrowsExceptionTest()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._009);

            // ACT
            var elements = pre.Slides[0].Shapes;

            pre.Close();
        }

        [Fact]
        public void PictureEx_BytesTest()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._009);
            var picEx = pre.Slides[1].Shapes.Single(e => e.Id.Equals(3));

            // ACT
            var bytes = picEx.Picture.ImageEx.Bytes;
            pre.Close();

            // ASSERT
            Assert.True(bytes.Length > 0);
        }

        [Fact]
        public void PictureEx_SetImageTest()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._009);
            var picEx = pre.Slides[1].Shapes.Single(e => e.Id.Equals(3));
            var testImage2Stream = new MemoryStream(Properties.Resources.test_image_2);
            var sizeBefore = picEx.Picture.ImageEx.Bytes.Length;

            // ACT
            picEx.Picture.ImageEx.SetImage(testImage2Stream);

            var sizeAfter = picEx.Picture.ImageEx.Bytes.Length;
            pre.Close();
            testImage2Stream.Dispose();

            // ASSERT
            Assert.NotEqual(sizeBefore,  sizeAfter);
        }

        [Fact]
        public void ShapeEx_BackgroundImage_BytesTest()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._009);
            var shapeEx = (ShapeSL)pre.Slides[2].Shapes.Single(e => e.Id.Equals(4));

            // ACT
            var length = shapeEx.BackgroundImage.Bytes.Length;

            // ASSERT
            Assert.True(length > 0);
        }

        [Fact]
        public void ShapeEx_BackgroundImage_SetImageTest()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._009);
            var shapeEx = (ShapeSL)pre.Slides[2].Shapes.Single(e => e.Id.Equals(4));
            var testImage2Stream = new MemoryStream(Properties.Resources.test_image_2);
            var sizeBefore = shapeEx.BackgroundImage.Bytes.Length;

            // ACT
            shapeEx.BackgroundImage.SetImage(testImage2Stream);

            var sizeAfter = shapeEx.BackgroundImage.Bytes.Length;
            pre.Close();
            testImage2Stream.Dispose();

            // ASSERT
            Assert.NotEqual(sizeBefore, sizeAfter);
        }

        [Fact]
        public void ShapeEx_BackgroundImage_IsNullTest()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._009);
            var shapeEx = (ShapeSL)pre.Slides[1].Shapes.Single(e => e.Id.Equals(6));

            // ACT
            var bImage = shapeEx.BackgroundImage;

            // ASSERT
            Assert.Null(bImage);
        }

        [Fact]
        public void OleObjects_ParseTest()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._009);
            var shapes = pre.Slides[1].Shapes;

            // ACT
            var oleNumbers = shapes.Count(e => e.Type.Equals(ShapeType.OLEObject));
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
            var pre = new PresentationSL(Properties.Resources._009);

            // ACT
            var name = pre.Slides[1].Shapes.Single(e => e.Id.Equals(8)).Name;

            // ASSERT
            Assert.Equal("Object 2", name);
        }

        [Fact]
        public void SlideEx_Background_IsNullTest()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._009);

            // ACT
            var bg = pre.Slides[1].BackgroundImage;

            // ASSERT
            Assert.Null(bg);
        }

        [Fact]
        public void SlideEx_Background_ChangeTest()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._009);
            var bg = pre.Slides[0].BackgroundImage;
            var testImage2Stream = new MemoryStream(Properties.Resources.test_image_2);
            var sizeBefore = bg.Bytes.Length;

            // ACT
            bg.SetImage(testImage2Stream);

            var sizeAfter = bg.Bytes.Length;
            pre.Close();
            testImage2Stream.Dispose();

            // ASSERT
            Assert.NotEqual(sizeBefore, sizeAfter);
        }

        [Fact]
        public void NumberParagraphAndPortionTest()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._009);
            var shape = (ShapeSL)pre.Slides[2].Shapes.SingleOrDefault(e => e.Id.Equals(2));
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
            var pre = new PresentationSL(Properties.Resources._009);

            // ACT
            var w = pre.SlideWidth;
            var y = pre.SlideHeight;

            pre.Close();

            // ASSERT
            Assert.Equal(9144000, w);
            Assert.Equal(5143500, y);
        }

        [Fact]
        public void TextBox_Placeholder_FontHeight_Case1_Test()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._009);
            var elements = pre.Slides[3].Shapes;
            var tb2TitlePh = elements.Single(e => e.Id.Equals(2));
            var tb3SubTitlePh = elements.Single(e => e.Id.Equals(3));

            // ACT
            var fhTitle = tb2TitlePh.TextFrame.Paragraphs.Single().Portions.Single().FontHeight;
            var text2 = tb2TitlePh.TextFrame.Text;
            var fhSubTitle = tb3SubTitlePh.TextFrame.Paragraphs.Single().Portions.Single().FontHeight;

            pre.Close();

            // ASSERT
            Assert.Equal(4400, fhTitle);
            Assert.Equal(3200, fhSubTitle);
            Assert.Equal("Title text", text2);
        }

        [Fact]
        public void TextBox_Placeholder_FontHeight_Case2_Test()
        {
            // ARRANGE
            var pre010 = new PresentationSL(Properties.Resources._010);
            var pre010TextBox = pre010.Slides.First().Shapes.First();

            // ACT
            var fh = pre010TextBox.TextFrame.Paragraphs.First().Portions.First().FontHeight;

            pre010.Close();

            // ASSERT
            Assert.Equal(1226, fh);
        }

        [Fact]
        public void TablesPropertiesTest()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._009);
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
        public void ChartPropertiesTest()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._009);
            var sld3Elements = pre.Slides[2].Shapes;
            var sld5Elements = pre.Slides[4].Shapes;
            var chartEx6 = sld3Elements.Single(e => e.Id.Equals(6));
            var chartEx7 = sld3Elements.Single(e => e.Id.Equals(7));
            var sld5Chart6 = sld5Elements.Single(e => e.Id.Equals(6));
            var sld5Chart3 = sld5Elements.Single(e => e.Id.Equals(3));
            var sld5Chart5 = sld5Elements.Single(e => e.Id.Equals(5));

            // ACT
            var chart7Title = chartEx7.Chart.Title;
            var chart6Title = chartEx6.Chart.Title;
            var chart7Type = chartEx7.Chart.Type;
            var sld5Chart6Title = sld5Chart6.Chart.Title;
            var sld5Chart3Title = sld5Chart3.Chart.Title;
            var sld5Chart5Title = sld5Chart5.Chart.Title;

            pre.Close();

            // ASSERT
            Assert.Equal("Sales", chart7Title);
            Assert.Equal("Sales2", chart6Title);
            Assert.Equal("Sales3", sld5Chart6Title);
            Assert.Equal("Sales4", sld5Chart3Title);
            Assert.Equal("Sales5", sld5Chart5Title);
            Assert.Equal(ChartType.PieChart, chart7Type);
        }
    }
}
