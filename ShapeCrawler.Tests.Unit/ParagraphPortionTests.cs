using System;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Tables;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Properties;
using Xunit;

// ReSharper disable All
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit
{
    public class ParagraphPortionTests : IClassFixture<PresentationFixture>
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
        public void Text_GetterThrowsElementIsRemovedException_WhenPortionIsRemoved()
        {
            // Arrange
            IPresentation presentation = SCPresentation.Open(Resources._001, true);
            IAutoShape autoShape = (IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 5);
            IPortionCollection portions = autoShape.TextBox.Paragraphs[0].Portions;
            IPortion portion = portions[0];
            portions.Remove(portion);

            // Act-Assert
            portion.Invoking(p => p.Text).Should().Throw<ElementIsRemovedException>();
        }
    }
}
