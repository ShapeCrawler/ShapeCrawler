using System.Linq;
using SlideDotNet.Enums;
using SlideDotNet.Models;
using Xunit;

// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.UnitTests
{
    public class TestFile_013
    {
        [Fact]
        public void ChartPropertiesTest()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._013);
            var combChart = pre.Slides[0].Shapes.Single(x => x.Id == 5).Chart;
            var chart4 = pre.Slides[0].Shapes.Single(x => x.Id == 4).Chart;

            // ACT
            var title = combChart.Title;
            var hasTitle = chart4.HasTitle;
            var type = combChart.Type;
            var numSeries = combChart.SeriesCollection.Count;

            pre.Close();

            // ASSERT
            Assert.Equal(ChartType.Combination, type);
            Assert.Equal("Title text", title);
            Assert.False(hasTitle);
            Assert.Equal(3, numSeries);
        }

        [Fact]
        public void Slide_Shapes_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._013);

            // ACT
            var shapes = pre.Slides[0].Shapes; // should not throw exception

            pre.Close();
        }

        [Fact]
        public void PlaceholderType_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._013);

            // ACT
            var phType = pre.Slides[0].Shapes.Single(s=>s.Id == 281).PlaceholderType;

            // ARRANGE
            Assert.Equal(PlaceholderType.Custom, phType);
        }

        [Fact]
        public void Chart_Title_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._013);
            var chart = pre.Slides[0].Shapes.Single(s => s.Id == 6).Chart;

            // ACT
            var hasTitle = chart.HasTitle;

            // ARRANGE
            Assert.False(hasTitle);
        }

        [Fact]
        public void TextFrame_Text_Test()
        {
            // ARRANGE
            var pre = new PresentationEx(Properties.Resources._014);
            var elId61 = pre.Slides[0].Shapes.Single(s => s.Id == 61);

            // ACT
            var text = elId61.TextFrame.Text;

            // ARRANGE
            Assert.NotNull(text);
        }
    }
}
