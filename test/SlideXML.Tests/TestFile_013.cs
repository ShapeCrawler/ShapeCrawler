using System.Linq;
using SlideXML.Enums;
using SlideXML.Models;
using Xunit;

namespace SlideXML.Tests
{
    public class TestFile_013
    {
        [Fact]
        public void ChartPropertiesTest()
        {
            // ARRANGE
            var pre13 = new PresentationSL(Properties.Resources._013);
            var chart = pre13.Slides[0].Shapes.Single(x => x.Id == 5).Chart;
            var chart4 = pre13.Slides[0].Shapes.Single(x => x.Id == 4).Chart;

            // ACT
            var title = chart.Title;
            var title4 = chart4.Title;
            var type = chart.Type;

            pre13.Close();

            // ASSERT
            Assert.Equal(ChartType.Combination, type);
            Assert.Equal("Title text", title);
            Assert.Null(title4);
        }

        [Fact]
        public void Slide_Shapes_Test()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._013);

            // ACT
            var shapes = pre.Slides[0].Shapes; // should not throw exception

            pre.Close();
        }

        [Fact]
        public void PlaceholderType_Test()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._013);

            // ACT
            var phType = pre.Slides[0].Shapes.Single(s=>s.Id == 281).PlaceholderType;

            // ARRANGE
            Assert.Equal(PlaceholderType.Custom, phType);
        }
    }
}
