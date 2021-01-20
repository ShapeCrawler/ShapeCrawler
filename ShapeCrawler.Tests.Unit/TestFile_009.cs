using System.IO;
using System.Linq;
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
        public void DateTimePlaceholder_Text_Test()
        {
            // ARRANGE
            var pre = new PresentationSc(Properties.Resources._011_dt);
            var dt = pre.Slides[0].Shapes.Single(s => s.Id == 54275);

            // ACT
            var hasText = dt.HasTextFrame;

            pre.Close();

            // ASSERT
            Assert.True(hasText);
        }

        [Fact]
        public void Placeholder_FontHeight_Test()
        {
            // ARRANGE
            var pre = new PresentationSc(Properties.Resources._012_title_placeholder);
            var title = pre.Slides[0].Shapes.Single(x => x.Id == 2);

            // ACT
            var text = title.Text.Content;
            var fh = title.Text.Paragraphs.First().Portions.First().Font.Size;

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
            var fh = pre010TextBox.Text.Paragraphs.First().Portions.First().Font.Size;

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
            var text = autoShape.Text.Content;

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
            var text = autoShape.Text.Content;
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