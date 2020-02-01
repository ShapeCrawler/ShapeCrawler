using System.Linq;
using SlideXML.Models;
using Xunit;

namespace SlideXML.Tests
{
    public class ShapeTextTests
    {
        [Fact]
        public void Shape_Text_Tests()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._011_dt);
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
        public void Shape_Text_Test2()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._012_title_placeholder);
            var autoShape = pre.Slides[0].Shapes.Single(s => s.Id == 3);

            // ACT
            var text = autoShape.TextFrame.Text;

            pre.Close();

            // ASSERT
            Assert.Equal("P1 P2", text);
        }
    }
}
