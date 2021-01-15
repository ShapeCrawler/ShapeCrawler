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
        public void Table_Row_Remove_Test()
        {
            // ARRANGE
            var pre = new PresentationSc(Properties.Resources._009);
            var sld3Shapes = pre.Slides[2].Shapes;
            var table3 = sld3Shapes.First(s => s.Id.Equals(3)).Table;
            var rows = table3.Rows;
            var numRowsBefore = rows.Count;

            // ACT
            rows.RemoveAt(0);
            
            var ms = new MemoryStream();
            pre.SaveAs(ms);
            pre.Close();

            pre = new PresentationSc(ms);
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
            var pre = new PresentationSc(Properties.Resources._011_dt);
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
            var pre = new PresentationSc(Properties.Resources._012_title_placeholder);
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
            var pre010 = new PresentationSc(Properties.Resources._010);
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
            var pre = new PresentationSc(Properties.Resources._006_1_slides);
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
            var pre = new PresentationSc(Properties.Resources._012_title_placeholder);
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
            var pre = new PresentationSc(Properties.Resources._011_dt);
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
            var pre = new PresentationSc(ms);

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
