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
        public void Remove_RemovesPortionFromCollection()
        {
            // Arrange
            var presentation = SCPresentation.Open(Resources._002, true);
            var portions = GetPortions(presentation);
            var portion = portions.First();
            var countBefore = portions.Count;

            // Act
            portion.Remove();

            // Assert
            portions.Should().HaveCount(1);
            portions.Should().HaveCountLessThan(countBefore);

            var memoryStream = new MemoryStream();
            presentation.SaveAs(memoryStream);
            var savedPresentation = SCPresentation.Open(memoryStream, false);
            portions = GetPortions(savedPresentation);
            portions.Should().HaveCount(1);
        }

        [Fact]
        public void Text_GetterReturnsParagraphPortionText()
        {
            // Arrange
            Portion portion = ((ITable)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0].TextBox
                .Paragraphs[0].Portions[0];

            // Act
            string paragraphPortionText = portion.Text;

            // Assert
            paragraphPortionText.Should().BeEquivalentTo("0:0_p1_lvl1");
        }

        private static PortionCollection GetPortions(SCPresentation presentation)
        {
            IAutoShape shape5 = presentation.Slides[1].Shapes.First(x => x.Id == 5) as IAutoShape;
            var portions = shape5.TextBox.Paragraphs[0].Portions;
            return portions;
        }

        
    }
}
