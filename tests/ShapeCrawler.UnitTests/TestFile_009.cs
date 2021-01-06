using System.IO;
using System.Linq;
using ShapeCrawler.Enums;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using Xunit;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit
{
    public class TestFile_009 : IClassFixture<TestFile_009Fixture>
    {
        private readonly TestFile_009Fixture _fixture;

        public TestFile_009(TestFile_009Fixture fixture)
        {
            _fixture = fixture;
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
            var bg = pre.Slides[1].Background;

            // ASSERT
            Assert.Null(bg);
        }

        [Fact]
        public void NumberParagraphAndPortionTest()
        {
            // ARRANGE
            var pre = _fixture.pre009;
            var shape = (Shape)pre.Slides[2].Shapes.SingleOrDefault(e => e.Id.Equals(2));
            var paragraphs = shape.TextFrame.Paragraphs;

            // ACT
            var numParagraphs = paragraphs.Count;
            var portions = paragraphs[0].Portions;
            var numPortions = portions.Count;
            var por1Size = portions[0].Font.Size;
            var por2Size = portions[1].Font.Size;


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
            var fhTitle = tb2TitlePh.TextFrame.Paragraphs.Single().Portions.Single().Font.Size;
            var text2 = tb2TitlePh.TextFrame.Text;
            var fhSubTitle = subTitle3.TextFrame.Paragraphs.Single().Portions.Single().Font.Size;

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
            var pre = new Presentation(Properties.Resources._009);
            var sld3Shapes = pre.Slides[2].Shapes;
            var table3 = sld3Shapes.First(s => s.Id.Equals(3)).Table;
            var rows = table3.Rows;
            var numRowsBefore = rows.Count;

            // ACT
            rows.RemoveAt(0);
            
            var ms = new MemoryStream();
            pre.SaveAs(ms);
            pre.Close();

            pre = new Presentation(ms);
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
            var pre = new Presentation(Properties.Resources._011_dt);
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
            var pre = new Presentation(Properties.Resources._012_title_placeholder);
            var title = pre.Slides[0].Shapes.Single(x => x.Id == 2);

            // ACT
            var text = title.TextFrame.Text;
            var fh = title.TextFrame.Paragraphs.First().Portions.First().Font.Size;

            pre.Close();

            // ASSERT
            Assert.NotNull(text);
            Assert.Equal(2000, fh);
        }

        [Fact]
        public void Placeholder_FontHeight_Title_Test()
        {
            // ARRANGE
            var pre010 = new Presentation(Properties.Resources._010);
            var pre010TextBox = pre010.Slides[0].Shapes.Single(x => x.Id == 2);

            // ACT
            var fh = pre010TextBox.TextFrame.Paragraphs.First().Portions.First().Font.Size;

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
            var pre = new Presentation(Properties.Resources._006_1_slides);
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
            var pre = new Presentation(Properties.Resources._012_title_placeholder);
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
            var pre = new Presentation(Properties.Resources._011_dt);
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
            var pre = new Presentation(ms);

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
    }
}
