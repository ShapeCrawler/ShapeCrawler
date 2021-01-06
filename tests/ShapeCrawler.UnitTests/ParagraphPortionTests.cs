using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Models;
using ShapeCrawler.Models.TextBody;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Properties;
using Xunit;

// ReSharper disable All
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit
{
    public class ParagraphPortionTests : IClassFixture<ReadOnlyTestPresentations>
    {
        private readonly ReadOnlyTestPresentations _fixture;

        public ParagraphPortionTests(ReadOnlyTestPresentations fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Remove_RemovesPortionFromCollection()
        {
            // Arrange
            var presentation = Presentation.Open(Resources._002, true);
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
            var savedPresentation = new Presentation(memoryStream, false);
            portions = GetPortions(savedPresentation);
            portions.Should().HaveCount(1);
        }

        [Fact]
        public void FontName_GetterReturnsFontNameOfTheParagraphPortion()
        {
            // Arrange
            ITextFrame textFrameCase1 = _fixture.Pre002.Slides[1].Shapes.First(s => s.Id == 3).TextFrame;
            ITextFrame textFrameCase2 = _fixture.Pre001.Slides[0].Shapes.First(s => s.Id == 4).TextFrame;

            // Act
            string portionFontNameCase1 = textFrameCase1.Paragraphs[0].Portions[0].Font.Name;
            string portionFontNameCase2 = textFrameCase2.Paragraphs[0].Portions[0].Font.Name;

            // Assert
            portionFontNameCase1.Should().BeEquivalentTo("Palatino Linotype");
            portionFontNameCase2.Should().BeEquivalentTo("Broadway");
        }

        [Fact]
        public void FontName_SetterSetsSpecifiedFontName()
        {
            // Arrange
            const string newFont = "Time New Roman";
            Portion paragraphPortion = Presentation.Open(Resources._001, true).
                Slides[0].Shapes.First(sp => sp.Id == 4).
                TextFrame.Paragraphs[0].Portions[0];
            // Act
            paragraphPortion.Font.Name = newFont;

            // Assert
            paragraphPortion.Font.Name.Should().BeEquivalentTo(newFont);
        }

        [Fact]
        public void FontName_SetterThrowsException_WhenAUserTrySetFontNameForPortionOfAPlaceholderShape()
        {
            // Arrange
            ITextFrame textFrame = Presentation.Open(Resources._001, true).Slides[2].Shapes
                .First(sp => sp.Id == 4).TextFrame;
            IList<Paragraph> paragraphs = textFrame.Paragraphs;
            Portion paragraphPortion = paragraphs[0].Portions[0];

            // Act
            Action action = () => paragraphPortion.Font.Name = "Time New Roman";

            // Assert
            action.Should().Throw<PlaceholderCannotBeChangedException>();
        }

        private static PortionCollection GetPortions(Presentation presentation)
        {
            var shape5 = presentation.Slides[1].Shapes.First(x => x.Id == 5);
            var portions = shape5.TextFrame.Paragraphs.First().Portions;
            return portions;
        }
    }
}
