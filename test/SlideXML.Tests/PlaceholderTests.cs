using System.IO;
using System.Linq;
using SlideXML.Enums;
using SlideXML.Models;
using Xunit;

namespace SlideXML.Tests
{
    /// <summary>
    /// Contains tests for placeholder shapes.
    /// </summary>
    public class PlaceholderShapeTest
    {
        [Fact]
        public void DateTimePlaceholder_HasTextFrame_Test()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._008);
            var sp3 = pre.Slides[0].Shapes.Single(sp => sp.Id == 3);

            // ACT
            var hasTextBody = sp3.HasTextFrame;

            pre.Close();

            // ASSERT
            Assert.False(hasTextBody);
        }

        [Fact]
        public void DateTimePlaceholder_Text_Test()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._011_dt);
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
        public void TitlePlaceholder_TextAndFont_Test()
        {
            // ARRANGE
            var pre = new PresentationSL(Properties.Resources._012_title_placeholder);
            var title = pre.Slides[0].Shapes.Single(x => x.Id == 2);

            // ACT
            var text = title.TextFrame.Text;
            var fh = title.TextFrame.Paragraphs.First().Portions.First().FontHeight;

            pre.Close();

            // ASSERT
            Assert.NotNull(text);
            Assert.Equal(2000, fh);
        }

        [Fact]
        public void TitlePlaceholder_FontHeight_Test()
        {
            // ARRANGE
            var pre010 = new PresentationSL(Properties.Resources._010);
            var pre010TextBox = pre010.Slides[0].Shapes.Single(x => x.Id == 2);

            // ACT
            var fh = pre010TextBox.TextFrame.Paragraphs.First().Portions.First().FontHeight;

            pre010.Close();

            // ASSERT
            Assert.Equal(1539, fh);
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
    }
}
