using System.Linq;
using FluentAssertions;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Tests.Helpers;
using Xunit;

// ReSharper disable All
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests
{
    public class ParagraphPortionTests : ShapeCrawlerTest, IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public ParagraphPortionTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Text_GetterReturnsParagraphPortionText()
        {
            // Arrange
            IPortion portion = ((ITable)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0].TextBox
                .Paragraphs[0].Portions[0];

            // Act
            string paragraphPortionText = portion.Text;

            // Assert
            paragraphPortionText.Should().BeEquivalentTo("0:0_p1_lvl1");
        }

        [Fact]
        public void Text_SetterThrowsElementIsRemovedException_WhenPortionIsRemoved()
        {
            // Arrange
            IPresentation presentation = SCPresentation.Open(TestFiles.Presentations.pre001, true);
            IAutoShape autoShape = (IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 5);
            IPortionCollection portions = autoShape.TextBox.Paragraphs[0].Portions;
            IPortion portion = portions[0];
            portions.Remove(portion);

            // Act-Assert
            portion.Invoking(p => p.Text = "new text").Should().Throw<ElementIsRemovedException>();
        }

        [Fact]
        public void Hyperlink_Setter_sets_hyperlink()
        {
            // Arrange
            var pptxStream = GetTestPptxStream("001.pptx");
            var presentation = SCPresentation.Open(pptxStream, true);
            var autoShape = presentation.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 3");
            var portion = autoShape.TextBox.Paragraphs[0].Portions[0];
            
            // Act
            portion.Hyperlink = "https://github.com/ShapeCrawler/ShapeCrawler";
            
            // Assert
            presentation.Save();
            presentation.Close();
            presentation = SCPresentation.Open(pptxStream, false);
            autoShape = presentation.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 3");
            portion = autoShape.TextBox.Paragraphs[0].Portions[0];
            portion.Hyperlink.Should().Be("https://github.com/ShapeCrawler/ShapeCrawler");
        }
    }
}
