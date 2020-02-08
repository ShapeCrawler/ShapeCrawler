using System.Linq;
using SlideXML.Models;
using Xunit;
// ReSharper disable TooManyChainedReferences

namespace SlideXML.Tests
{
    public class TestFile_014
    {
        [Fact]
        public void TextFrame_Text_Test()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._014);
            var elId61 = pre.Slides[0].Elements.Single(s => s.Id == 61);

            // ACT
            var text = elId61.TextFrame.Text;

            // ARRANGE
            Assert.NotNull(text);
        }

        [Fact]
        public void Portion_FontHeight_Test()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._014);
            var elId5 = pre.Slides[1].Elements.Single(x => x.Id == 5);

            // ACT-ASSERT
            var text = elId5.TextFrame.Text;
            var fh = elId5.TextFrame.Paragraphs.First().Portions.First().FontHeight;
        }

        [Fact]
        public void Slide_Elements_Test()
        {
            // ARRANGE
            var pre = new Presentation(Properties.Resources._014);

            // ACT-ASSERT
            var elements = pre.Slides[2].Elements;
        }
    }
}
