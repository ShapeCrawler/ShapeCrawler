using System.Linq;
using ShapeCrawler.Enums;
using ShapeCrawler.Models;
using Xunit;

// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.Tests.Unit
{
    public class TestFile_013
    {


        [Fact]
        public void PlaceholderType_Test()
        {
            // ARRANGE
            var pre = new PresentationSc(Properties.Resources._013);

            // ACT
            var phType = pre.Slides[0].Shapes.Single(s=>s.Id == 281).PlaceholderType;

            // ARRANGE
            Assert.Equal(PlaceholderType.Custom, phType);
        }

        [Fact]
        public void Chart_Title_Test()
        {
            // ARRANGE
            var pre = new PresentationSc(Properties.Resources._013);
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
            var pre = new PresentationSc(Properties.Resources._014);
            var elId61 = pre.Slides[0].Shapes.Single(s => s.Id == 61);

            // ACT
            var text = elId61.TextFrame.Text;

            // ARRANGE
            Assert.NotNull(text);
        }
    }
}
