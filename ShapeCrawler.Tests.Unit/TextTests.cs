using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Enums;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Properties;
using ShapeCrawler.Texts;
using Xunit;

// ReSharper disable All
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit
{
    public class TextTests : IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public TextTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Content_GetterReturnsShapeTextWhichIsParagraphsTextAggregate()
        {
            // Arrange
            TextSc textFrameCase1 = _fixture.Pre009.Slides[3].Shapes.First(sp => sp.Id == 2).Text;
            TextSc textFrameCase2 = _fixture.Pre001.Slides[0].Shapes.First(sp => sp.Id == 5).Text;
            TextSc textFrameCase3 = _fixture.Pre001.Slides[0].Shapes.First(sp => sp.Id == 6).Text;
            TextSc textFrameCase4 = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3).Table.Rows[0].Cells[0].Text;
            TextSc textFrameCase5 = _fixture.Pre019.Slides[0].Shapes.First(sp => sp.Id == 2).Text;
            TextSc textFrameCase6 = _fixture.Pre014.Slides[0].Shapes.First(sp => sp.Id == 61).Text;
            TextSc textFrameCase7 = _fixture.Pre014.Slides[1].Shapes.First(sp => sp.Id == 5).Text;

            // Act
            string shapeTextCase1 = textFrameCase1.Content;
            string shapeTextCase2 = textFrameCase2.Content;
            string shapeTextCase3 = textFrameCase3.Content;
            string shapeTextCase4 = textFrameCase4.Content;
            string shapeTextCase5 = textFrameCase5.Content;
            string shapeTextCase6 = textFrameCase6.Content;
            string shapeTextCase7 = textFrameCase7.Content;

            // Assert
            shapeTextCase1.Should().BeEquivalentTo("Title text");
            shapeTextCase2.Should().BeEquivalentTo(" id5-Text1");
            shapeTextCase3.Should().BeEquivalentTo($"id6-Text1{Environment.NewLine}Text2");
            shapeTextCase4.Should().BeEquivalentTo($"0:0_p1_lvl1{Environment.NewLine}0:0_p2_lvl2");
            shapeTextCase5.Should().BeEquivalentTo("1");
            shapeTextCase6.Should().BeEquivalentTo($"test1{Environment.NewLine}test2{Environment.NewLine}" +
                                                   $"test3{Environment.NewLine}test4{Environment.NewLine}test5");
            shapeTextCase7.Should().BeEquivalentTo($"Test subtitle");
        }

        [Fact]
        public void Content_SetterChangesTextByUsingFirstParagraphAsBasicSingleParagraph()
        {
            // Arrange
            PresentationSc presentation = PresentationSc.Open(Resources._001, true);
            TextSc text = presentation.Slides[0].Shapes.First(sp => sp.Id == 3).Text;
            const string newText = "Test";
            var mStream = new MemoryStream();

            // Act
            text.Content = newText;

            // Assert
            text.Content.Should().BeEquivalentTo(newText);
            text.Paragraphs.Should().HaveCount(1);
            
            presentation.SaveAs(mStream);
            presentation.Close();
            presentation = PresentationSc.Open(mStream, false);
            text = presentation.Slides[0].Shapes.First(sp => sp.Id == 3).Text;
            text.Content.Should().BeEquivalentTo(newText);
            text.Paragraphs.Should().HaveCount(1);
        }

        [Fact]
        public void HasTextFrameAndText_PropertiesReturnCorrectValue_WhenTheirGettersAreCalled()
        {
            // Arrange
            var shapeCase1 = _fixture.Pre008.Slides[0].Shapes.First(sp => sp.Id == 3);
            var shapeCase2 = _fixture.Pre021.Slides[3].Shapes.First(sp => sp.Id == 2);
            var textFrameCase1 = shapeCase1.Text;
            var textFrameCase2 = shapeCase2.Text;

            // Act
            var hasTextFrameCase1 = shapeCase1.HasTextFrame;
            var hasTextFrameCase2 = shapeCase2.HasTextFrame;
            var frameTextCase1 = textFrameCase1.Content;
            var frameTextCase2 = textFrameCase2.Content;

            // Assert
            hasTextFrameCase1.Should().BeTrue();
            hasTextFrameCase2.Should().BeTrue();
            frameTextCase1.Should().BeEquivalentTo("25.01.2020");
            frameTextCase2.Should().BeEquivalentTo("test footer");
        }

        [Fact]
        public void ParagraphBulletFontNameProperty_ReturnsFontName()
        {
            // Arrange
            var shapes = _fixture.Pre002.Slides[1].Shapes;
            var shape3Pr1Bullet = shapes.First(x => x.Id == 3).Text.Paragraphs[0].Bullet;
            var shape4Pr2Bullet = shapes.First(x => x.Id == 4).Text.Paragraphs[1].Bullet;

            // Act
            var shape3BulletFontName = shape3Pr1Bullet.FontName;
            var shape4BulletFontName = shape4Pr2Bullet.FontName;

            // Assert
            shape3BulletFontName.Should().BeNull();
            shape4BulletFontName.Should().Be("Calibri");
        }

        [Fact]
        public void ParagraphBulletTypeProperty_ReturnsCorrectValue()
        {
            // Arrange
            var shapeList = _fixture.Pre002.Slides[1].Shapes;
            var shape4 = shapeList.First(x => x.Id == 4);
            var shape5 = shapeList.First(x => x.Id == 5);
            var shape4Pr2Bullet = shape4.Text.Paragraphs[1].Bullet;
            var shape5Pr1Bullet = shape5.Text.Paragraphs[0].Bullet;
            var shape5Pr2Bullet = shape5.Text.Paragraphs[1].Bullet;

            // Act
            var shape5Pr1BulletType = shape5Pr1Bullet.Type;
            var shape5Pr2BulletType = shape5Pr2Bullet.Type;
            var shape4Pr2BulletType = shape4Pr2Bullet.Type;

            // Assert
            shape5Pr1BulletType.Should().BeEquivalentTo(BulletType.Numbered);
            shape5Pr2BulletType.Should().BeEquivalentTo(BulletType.Picture);
            shape4Pr2BulletType.Should().BeEquivalentTo(BulletType.Character);
        }

        [Fact]
        public void ParagraphBulletColorHexAndCharAndSizeProperties_ReturnCorrectValues()
        {
            // Arrange
            var shapeList = _fixture.Pre002.Slides[1].Shapes;
            var shape4 = shapeList.First(x => x.Id == 4);
            var shape4Pr2Bullet = shape4.Text.Paragraphs[1].Bullet;

            // Act
            var bulletColorHex = shape4Pr2Bullet.ColorHex;
            var bulletChar = shape4Pr2Bullet.Char;
            var bulletSize = shape4Pr2Bullet.Size;

            // Assert
            bulletColorHex.Should().Be("C00000");
            bulletChar.Should().Be("'");
            bulletSize.Should().Be(120);
        }

        [Theory]
        [MemberData(nameof(TestCasesParagraphText))]
        public void ParagraphText_SetterChangesParagraphText(
            PresentationSc presentation, 
            ElementRequest prRequest, 
            string newPrText,
            int expectedNumPortions)
        {
            // Arrange
            ParagraphSc paragraph = TestHelper.GetParagraph(presentation, prRequest);
            var presentationStream = new MemoryStream();

            // Act
            paragraph.Text = newPrText;

            // Assert
            paragraph.Text.Should().BeEquivalentTo(newPrText);
            paragraph.Portions.Should().HaveCount(expectedNumPortions);

            presentation.SaveAs(presentationStream);
            presentation.Close();
            paragraph = TestHelper.GetParagraph(presentationStream, prRequest);
            paragraph.Text.Should().BeEquivalentTo(newPrText);
            paragraph.Portions.Should().HaveCount(expectedNumPortions);
        }

        public static IEnumerable<object[]> TestCasesParagraphText()
        {
            var paragraphRequest = new ElementRequest
            {
                SlideIndex = 1,
                ShapeId = 4,
                ParagraphIndex = 1
            };
            PresentationSc presentation;
            paragraphRequest.ParagraphIndex = 2;

            presentation = PresentationSc.Open(Resources._002, true);
            yield return new object[] { presentation, paragraphRequest, "Text", 1};

            presentation = PresentationSc.Open(Resources._002, true);
            yield return new object[] { presentation, paragraphRequest, $"Text{Environment.NewLine}", 1};

            presentation = PresentationSc.Open(Resources._002, true);
            yield return new object[] { presentation, paragraphRequest, $"Text{Environment.NewLine}Text2", 2};

            presentation = PresentationSc.Open(Resources._002, true);
            yield return new object[] { presentation, paragraphRequest, $"Text{Environment.NewLine}Text2{Environment.NewLine}", 2 };
        }

        [Fact]
        public void ParagraphText_GetterReturnsParagraphText()
        {
            // Arrange
            TextSc textFrameCase1 = _fixture.Pre008.Slides[0].Shapes.First(sp => sp.Id == 37).Text;
            TextSc textFrameCase2 = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3).Table.Rows[0].Cells[0].Text;
            TextSc textFrameCase3 = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3).Table.Rows[0].Cells[0].Text;

            // Act
            string paragraphTextCase1 = textFrameCase1.Paragraphs[0].Text;
            string paragraphTextCase2 = textFrameCase1.Paragraphs[1].Text;
            string paragraphTextCase3 = textFrameCase2.Paragraphs[0].Text;

            // Assert
            paragraphTextCase1.Should().BeEquivalentTo("P1t1 P1t2");
            paragraphTextCase2.Should().BeEquivalentTo("p2");
            paragraphTextCase3.Should().BeEquivalentTo("0:0_p1_lvl1");
        }

        [Fact]
        public void Paragraphs_CollectionCounterReturnsNumberOfParagraphsInTheTextFrame()
        {
            // Arrange
            TextSc textFrame = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 2).Text;

            // Act
            IEnumerable<ParagraphSc> paragraphs = textFrame.Paragraphs;

            // Assert
            paragraphs.Should().HaveCount(1);
        }

        [Fact]
        public void ParagraphPortions_CollectionCounterReturnsNumberOfTextPortionsInTheParagraph()
        {
            // Arrange
            TextSc textFrame = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 2).Text;

            // Act
            IEnumerable<Portion> paragraphPortions = textFrame.Paragraphs[0].Portions;

            // Assert
            paragraphPortions.Should().HaveCount(2);
        }

        [Fact]
        public void ParagraphsCount_ReturnsNumberOfParagraphsInTheTextFrame()
        {
            // Arrange
            TextSc textFrame = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3).Table.Rows[0].Cells[0]
                .Text;

            // Act
            int paragraphsCount = textFrame.Paragraphs.Count;

            // Assert
            paragraphsCount.Should().Be(2);
        }
    }
}
