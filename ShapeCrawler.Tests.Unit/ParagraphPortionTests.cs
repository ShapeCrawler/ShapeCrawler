using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Models;
using ShapeCrawler.Models.TextShape;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Properties;
using Xunit;

// ReSharper disable All
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit
{
    public class ParagraphPortionTests : IClassFixture<PptxFixture>
    {
        private readonly PptxFixture _fixture;

        public ParagraphPortionTests(PptxFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Remove_RemovesPortionFromCollection()
        {
            // Arrange
            var presentation = PresentationEx.Open(Resources._002, true);
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
            var savedPresentation = new PresentationEx(memoryStream, false);
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
            Portion paragraphPortion = PresentationEx.Open(Resources._001, true).
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
            ITextFrame textFrame = PresentationEx.Open(Resources._001, true).Slides[2].Shapes
                .First(sp => sp.Id == 4).TextFrame;
            IList<ParagraphEx> paragraphs = textFrame.Paragraphs;
            Portion paragraphPortion = paragraphs[0].Portions[0];

            // Act
            Action action = () => paragraphPortion.Font.Name = "Time New Roman";

            // Assert
            action.Should().Throw<PlaceholderCannotBeChangedException>();
        }

        [Fact]
        public void FontSize_GetterReturnsFontSizeOfTheParagraphPortion()
        {
            // Arrange
            Portion portionCase1 = _fixture.Pre020.Slides[0].Shapes.First(sp => sp.Id == 3).
                                                                TextFrame.Paragraphs[0].Portions[0];
            Portion portionCase2 = _fixture.Pre015.Slides[0].Shapes.First(sp => sp.Id == 5).
                                                                TextFrame.Paragraphs[0].Portions[2];
            Portion portionCase3 = _fixture.Pre015.Slides[1].Shapes.First(sp => sp.Id == 61).
                                                                TextFrame.Paragraphs[0].Portions[0];
            Portion portionCase4 = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 2).
                                                                TextFrame.Paragraphs[0].Portions[0];
            Portion portionCase5 = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 2).
                                                                TextFrame.Paragraphs[0].Portions[1];
            Portion portionCase6 = _fixture.Pre009.Slides[3].Shapes.First(sp => sp.Id == 2).
                                                                TextFrame.Paragraphs[0].Portions[0];
            Portion portionCase7 = _fixture.Pre009.Slides[3].Shapes.First(sp => sp.Id == 3).
                                                                TextFrame.Paragraphs[0].Portions[0];

            // Act
            int portionFontSizeCase1 = portionCase1.Font.Size;
            int portionFontSizeCase2 = portionCase2.Font.Size;
            int portionFontSizeCase3 = portionCase3.Font.Size;
            int portionFontSizeCase4 = portionCase4.Font.Size;
            int portionFontSizeCase5 = portionCase5.Font.Size;
            int portionFontSizeCase6 = portionCase6.Font.Size;
            int portionFontSizeCase7 = portionCase7.Font.Size;

            // Assert
            portionFontSizeCase1.Should().Be(1800);
            portionFontSizeCase2.Should().Be(1800);
            portionFontSizeCase3.Should().Be(1867);
            portionFontSizeCase4.Should().Be(1800);
            portionFontSizeCase5.Should().Be(2000);
            portionFontSizeCase6.Should().Be(4400);
            portionFontSizeCase7.Should().Be(3200);
        }

        [Fact]
        public void FontSize_SetterChangesFontSizeOfParagraphPortion()
        {
            // Arrange
            static Portion GetPortion(PresentationEx presentation)
            {
                Portion portion = presentation.Slides[0].Shapes.First(sp => sp.Id == 4).TextFrame.Paragraphs[0].Portions[0];
                return portion;
            }
            int newFontSize = 28;
            var savedPreStream = new MemoryStream();
            PresentationEx presentation = PresentationEx.Open(Resources._001, true);
            Portion portion = GetPortion(presentation);
            int oldFontSize = portion.Font.Size;

            // Act
            portion.Font.Size = newFontSize;
            presentation.SaveAs(savedPreStream);

            // Assert
            presentation = PresentationEx.Open(savedPreStream, false);
            portion = GetPortion(presentation);
            portion.Font.Size.Should().NotBe(oldFontSize);
            portion.Font.Size.Should().Be(newFontSize);
            portion.Font.SizeCanBeChanged().Should().BeTrue();
        }

        private static PortionCollection GetPortions(PresentationEx presentation)
        {
            var shape5 = presentation.Slides[1].Shapes.First(x => x.Id == 5);
            var portions = shape5.TextFrame.Paragraphs.First().Portions;
            return portions;
        }
    }
}
