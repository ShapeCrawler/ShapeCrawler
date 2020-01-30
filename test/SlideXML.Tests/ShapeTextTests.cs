using System.Linq;
using SlideXML.Enums;
using SlideXML.Models;
using Xunit;

namespace SlideXML.Tests
{
    public class ShapeTextTests
    {
        [Fact]
        public void AutoShape_Text_Tests()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._011_dt);
            var autoShape = pre.Slides[0].Shapes.Single(s=>s.Id == 2);

            // ACT
            var text = autoShape.TextFrame.Text;

            pre.Close();

            // ASSERT
            Assert.NotNull(text);
        }
    }
}
