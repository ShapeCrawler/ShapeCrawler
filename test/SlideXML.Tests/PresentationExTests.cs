using System.IO;
using System.Linq;
using SlideXML.Enums;
using SlideXML.Models;
using SlideXML.Models.Elements;
using Xunit;

namespace SlideXML.Tests
{
    /// <summary>
    /// Represents tests of the <see cref="PresentationEx"/> class.
    /// </summary>
    public class PresentationExTests
    {
        [Fact]
        public void SlidesNumber_Test()
        {
            var ms = new MemoryStream(Properties.Resources._001);
            var pre = new PresentationEx(ms);

            // ACT
            var sldNumber = pre.Slides.Count();

            // CLOSE
            pre.Dispose();
            
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
            pre.Dispose();

            var pre2 = new PresentationEx(ms);
            var numSlides = pre2.Slides.Count();
            var numElements = pre2.Slides.Single().Elements.Count;
            pre2.Dispose();
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
            var shapes = pre.Slides.Single().Elements.OfType<ShapeEx>();
            var sh36 = shapes.Single(e => e.Id == 36);
            var sh37 = shapes.Single(e => e.Id == 37);
            pre.Dispose();

            // ASSERT
            Assert.Null(sh36.TextBody);
            Assert.NotNull(sh37.TextBody);
            Assert.Equal("P1t1 P1t2", sh37.TextBody.Paragraphs[0].Text);
            Assert.Equal("p2", sh37.TextBody.Paragraphs[1].Text);
        }

        [Fact]
        public void SlideElementsCount_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._003);

            // ACT
            var numberElements = pre.Slides.Single().Elements.Count;
            pre.Dispose();

            // ASSERT
            Assert.Equal(5, numberElements);
        }

        [Fact]
        public void ShapePlaceholderTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._006_1_slides);

            // ACT
            var shapePlaceholder = pre.Slides.Single().Elements.Single();
            pre.Dispose();

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
            var pre = new PresentationEx(Properties.Resources._009);

            // ACT
            var slides = pre.Slides;
            var groupElement = (GroupEx)pre.Slides[1].Elements.Single(x => x.Type.Equals(ElementType.Group));
            var el3 = groupElement.Elements.Single(x => x.Id.Equals(5));
            pre.Dispose();

            // ASSERT
            Assert.Equal(1581846, el3.X);
            Assert.Equal(1181377, el3.Width);
            Assert.Equal(654096, el3.Height);
        }

        [Fact]
        public void SecondSlideElementsNumberTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);

            // ACT
            var elNumber1 = pre.Slides[0].Elements.Count;
            var elNumber2 = pre.Slides[1].Elements.Count;
            pre.Dispose();

            // ASSERT
            Assert.Equal(6, elNumber1);
            Assert.Equal(5, elNumber2);
        }

        [Fact]
        public void SlideElementsDoNotThrowsExceptionTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);

            // ACT
            var elements = pre.Slides[0].Elements;

            pre.Dispose();
        }

        [Fact]
        public void PictureEx_BytesTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
            var picEx = (PictureEx) pre.Slides[1].Elements.Single(e => e.Id.Equals(3));

            // ACT
            var bytes = picEx.ImageEx.Bytes;
            pre.Dispose();

            // ASSERT
            Assert.True(bytes.Length > 0);
        }

        [Fact]
        public void PictureEx_SetImageTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
            var picEx = (PictureEx)pre.Slides[1].Elements.Single(e => e.Id.Equals(3));
            var testImage2Stream = new MemoryStream(Properties.Resources.test_image_2);
            var sizeBefore = picEx.ImageEx.Bytes.Length;

            // ACT
            picEx.ImageEx.SetImage(testImage2Stream);

            var sizeAfter = picEx.ImageEx.Bytes.Length;
            pre.Dispose();
            testImage2Stream.Dispose();

            // ASSERT
            Assert.NotEqual(sizeBefore,  sizeAfter);
        }

        [Fact]
        public void ShapeEx_BackgroundImage_BytesTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
            var shapeEx = (ShapeEx)pre.Slides[2].Elements.Single(e => e.Id.Equals(4));

            // ACT
            var length = shapeEx.BackgroundImage.Bytes.Length;

            // ASSERT
            Assert.True(length > 0);
        }

        [Fact]
        public void ShapeEx_BackgroundImage_SetImageTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
            var shapeEx = (ShapeEx)pre.Slides[2].Elements.Single(e => e.Id.Equals(4));
            var testImage2Stream = new MemoryStream(Properties.Resources.test_image_2);
            var sizeBefore = shapeEx.BackgroundImage.Bytes.Length;

            // ACT
            shapeEx.BackgroundImage.SetImage(testImage2Stream);

            var sizeAfter = shapeEx.BackgroundImage.Bytes.Length;
            pre.Dispose();
            testImage2Stream.Dispose();

            // ASSERT
            Assert.NotEqual(sizeBefore, sizeAfter);
        }

        [Fact]
        public void ShapeEx_BackgroundImage_IsNullTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
            var shapeEx = (ShapeEx)pre.Slides[1].Elements.Single(e => e.Id.Equals(6));

            // ACT
            var bImage = shapeEx.BackgroundImage;

            // ASSERT
            Assert.Null(bImage);
        }

        [Fact]
        public void OLEObjects_ParseTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);

            // ACT
            var oleNumbers = pre.Slides[1].Elements.Count(e => e.Type.Equals(ElementType.OLEObject));

            // ASSERT
            Assert.Equal(2, oleNumbers);
        }

        [Fact]
        public void OLEObject_NameTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);

            // ACT
            var name = pre.Slides[1].Elements.Single(e => e.Id.Equals(8)).Name;

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
            var sizeBefore = bg.Bytes.Length;

            // ACT
            bg.SetImage(testImage2Stream);

            var sizeAfter = bg.Bytes.Length;
            pre.Dispose();
            testImage2Stream.Dispose();

            // ASSERT
            Assert.NotEqual(sizeBefore, sizeAfter);
        }

        [Fact]
        public void NumberParagraphAndPortionTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
            var shape = (ShapeEx)pre.Slides[2].Elements.SingleOrDefault(e => e.Id.Equals(2));
            var paragraphs = shape.TextBody.Paragraphs;

            // ACT
            var numParagraphs = paragraphs.Count;
            var portions = paragraphs[0].Portions;
            var numPortions = portions.Count;
            var por1Size = portions[0].FontHeight;
            var por2Size = portions[1].FontHeight;

            pre.Dispose();

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

            pre.Dispose();

            // ASSERT
            Assert.Equal(9144000, w);
            Assert.Equal(5143500, y);
        }

        [Fact]
        public void Placeholder_FontHeightTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
            var elements = pre.Slides[3].Elements;
            var titlePlaceholder = (ShapeEx)elements.Single(e => e.Id.Equals(2));
            var subTitlePlaceholder = (ShapeEx)elements.Single(e => e.Id.Equals(3));

            // ACT
            var fhTitle = titlePlaceholder.TextBody.Paragraphs.Single().Portions.Single().FontHeight;
            var fhSubTitle = subTitlePlaceholder.TextBody.Paragraphs.Single().Portions.Single().FontHeight;

            pre.Dispose();

            // ASSERT
            Assert.Equal(4400, fhTitle);
            Assert.Equal(3200, fhSubTitle);
        }

        [Fact]
        public void TablesPropertiesTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._009);
            var elements = pre.Slides[2].Elements;
            var tblEx = (TableEx)elements.Single(e => e.Id.Equals(3));
            var firstRow = tblEx.Rows.First();

            // ACT
            var numRows = tblEx.Rows.Count;
            var numCells = firstRow.Cells.Count;
            var numParagraphs = firstRow.Cells.First().TextBody.Paragraphs.Count;
            var cellTxt = firstRow.Cells.First().TextBody.Text;
            var prText = firstRow.Cells.First().TextBody.Paragraphs.First().Text;
            var portionTxt = firstRow.Cells.First().TextBody.Paragraphs.First().Portions.Single().Text;

            pre.Dispose();

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
            var pre = new PresentationEx(Properties.Resources._009);
            var sld3Elements = pre.Slides[2].Elements;
            var sld5Elements = pre.Slides[4].Elements;
            var chartEx6 = (ChartEx)sld3Elements.Single(e => e.Id.Equals(6));
            var chartEx7 = (ChartEx)sld3Elements.Single(e => e.Id.Equals(7));
            var sld5Chart6 = (ChartEx)sld5Elements.Single(e => e.Id.Equals(6));
            var sld5Chart3 = (ChartEx)sld5Elements.Single(e => e.Id.Equals(3));
            var sld5Chart5 = (ChartEx)sld5Elements.Single(e => e.Id.Equals(5));

            // ACT
            var chart7Title = chartEx7.Title;
            var chart6Title = chartEx6.Title;
            var chart7Type = chartEx7.Type;
            var sld5Chart6Title = sld5Chart6.Title;
            var sld5Chart3Title = sld5Chart3.Title;
            var sld5Chart5Title = sld5Chart5.Title;

            pre.Dispose();

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
