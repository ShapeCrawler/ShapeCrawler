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
    public class TextBoxTests : IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public TextBoxTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Text_GetterReturnsShapeTextWhichIsParagraphTextsAggregate()
        {
            // Arrange
            TextBoxSc textCase1 = _fixture.Pre009.Slides[3].Shapes.First(sp => sp.Id == 2).TextBox;
            TextBoxSc textCase2 = _fixture.Pre001.Slides[0].Shapes.First(sp => sp.Id == 5).TextBox;
            TextBoxSc textCase3 = _fixture.Pre001.Slides[0].Shapes.First(sp => sp.Id == 6).TextBox;
            TextBoxSc textCase4 = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3).Table.Rows[0].Cells[0].TextBoxBox;
            TextBoxSc textCase5 = _fixture.Pre019.Slides[0].Shapes.First(sp => sp.Id == 2).TextBox;
            TextBoxSc textCase6 = _fixture.Pre014.Slides[0].Shapes.First(sp => sp.Id == 61).TextBox;
            TextBoxSc textCase7 = _fixture.Pre014.Slides[1].Shapes.First(sp => sp.Id == 5).TextBox;
            TextBoxSc textCase8 = _fixture.Pre011.Slides[0].Shapes.First(sp => sp.Id == 54275).TextBox;
            TextBoxSc textCase9 = _fixture.Pre008.Slides[0].Shapes.First(sp => sp.Id == 3).TextBox;
            TextBoxSc textCase10 = _fixture.Pre021.Slides[3].Shapes.First(sp => sp.Id == 2).TextBox;
            TextBoxSc textCase11 = _fixture.Pre012.Slides[0].Shapes.First(sp => sp.Id == 2).TextBox;

            // Act
            string textContentCase1 = textCase1.Text;
            string textContentCase2 = textCase2.Text;
            string textContentCase3 = textCase3.Text;
            string textContentCase4 = textCase4.Text;
            string textContentCase5 = textCase5.Text;
            string textContentCase6 = textCase6.Text;
            string textContentCase7 = textCase7.Text;
            string textContentCase8 = textCase8.Text;
            string textContentCase9 = textCase9.Text;
            string textContentCase10 = textCase10.Text;
            string textContentCase11 = textCase11.Text;

            // Assert
            textContentCase1.Should().BeEquivalentTo("Title text");
            textContentCase2.Should().BeEquivalentTo(" id5-Text1");
            textContentCase3.Should().BeEquivalentTo($"id6-Text1{Environment.NewLine}Text2");
            textContentCase4.Should().BeEquivalentTo($"0:0_p1_lvl1{Environment.NewLine}0:0_p2_lvl2");
            textContentCase5.Should().BeEquivalentTo("1");
            textContentCase6.Should().BeEquivalentTo($"test1{Environment.NewLine}test2{Environment.NewLine}" +
                                                   $"test3{Environment.NewLine}test4{Environment.NewLine}test5");
            textContentCase7.Should().BeEquivalentTo("Test subtitle");
            textContentCase8.Should().BeEquivalentTo("Jan 2018");
            textContentCase9.Should().BeEquivalentTo("25.01.2020");
            textContentCase10.Should().BeEquivalentTo("test footer");
            textContentCase11.Should().BeEquivalentTo("Test title text");
        }

        [Fact]
        public void Text_SetterChangesTextByUsingFirstParagraphAsBasicSingleParagraph()
        {
            // Arrange
            PresentationSc presentation = PresentationSc.Open(Resources._001, true);
            TextBoxSc text = presentation.Slides[0].Shapes.First(sp => sp.Id == 3).TextBox;
            const string newText = "Test";
            var mStream = new MemoryStream();

            // Act
            text.Text = newText;

            // Assert
            text.Text.Should().BeEquivalentTo(newText);
            text.Paragraphs.Should().HaveCount(1);
            
            presentation.SaveAs(mStream);
            presentation.Close();
            presentation = PresentationSc.Open(mStream, false);
            text = presentation.Slides[0].Shapes.First(sp => sp.Id == 3).TextBox;
            text.Text.Should().BeEquivalentTo(newText);
            text.Paragraphs.Should().HaveCount(1);
        }

        [Fact]
        public void HasTextBox_ReturnsTrue_WhenTheShapeContainsATextBox()
        {
            // Arrange
            ShapeSc shapeCase1 = _fixture.Pre008.Slides[0].Shapes.First(sp => sp.Id == 3);
            ShapeSc shapeCase2 = _fixture.Pre021.Slides[3].Shapes.First(sp => sp.Id == 2);
            ShapeSc shapeCase3 = _fixture.Pre011.Slides[0].Shapes.First(sp => sp.Id == 54275);

            // Act
            bool hasTextBoxCase1 = shapeCase1.HasTextBox;
            bool hasTextBoxCase2 = shapeCase2.HasTextBox;
            bool hasTextBoxCase3 = shapeCase3.HasTextBox;

            // Assert
            hasTextBoxCase1.Should().BeTrue();
            hasTextBoxCase2.Should().BeTrue();
            hasTextBoxCase3.Should().BeTrue();
        }

        [Fact]
        public void ParagraphBulletFontNameProperty_ReturnsFontName()
        {
            // Arrange
            var shapes = _fixture.Pre002.Slides[1].Shapes;
            var shape3Pr1Bullet = shapes.First(x => x.Id == 3).TextBox.Paragraphs[0].Bullet;
            var shape4Pr2Bullet = shapes.First(x => x.Id == 4).TextBox.Paragraphs[1].Bullet;

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
            var shape4Pr2Bullet = shape4.TextBox.Paragraphs[1].Bullet;
            var shape5Pr1Bullet = shape5.TextBox.Paragraphs[0].Bullet;
            var shape5Pr2Bullet = shape5.TextBox.Paragraphs[1].Bullet;

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
            var shape4Pr2Bullet = shape4.TextBox.Paragraphs[1].Bullet;

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
            TextBoxSc textFrameCase1 = _fixture.Pre008.Slides[0].Shapes.First(sp => sp.Id == 37).TextBox;
            TextBoxSc textFrameCase2 = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3).Table.Rows[0].Cells[0].TextBoxBox;
            TextBoxSc textFrameCase3 = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3).Table.Rows[0].Cells[0].TextBoxBox;

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
            TextBoxSc textFrame = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 2).TextBox;

            // Act
            IEnumerable<ParagraphSc> paragraphs = textFrame.Paragraphs;

            // Assert
            paragraphs.Should().HaveCount(1);
        }

        [Fact]
        public void ParagraphPortions_CollectionCounterReturnsNumberOfTextPortionsInTheParagraph()
        {
            // Arrange
            TextBoxSc textFrame = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 2).TextBox;

            // Act
            IEnumerable<Portion> paragraphPortions = textFrame.Paragraphs[0].Portions;

            // Assert
            paragraphPortions.Should().HaveCount(2);
        }

        [Fact]
        public void ParagraphsCount_ReturnsNumberOfParagraphsInTheTextFrame()
        {
            // Arrange
            TextBoxSc textFrame = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3).Table.Rows[0].Cells[0]
                .TextBoxBox;

            // Act
            int paragraphsCount = textFrame.Paragraphs.Count;

            // Assert
            paragraphsCount.Should().Be(2);
        }

        [Fact]
        public void ParagraphsAdd_AddsANewTextParagraphAtTheEndOfTheTextBoxAndReturnsAddedParagraph()
        {
            // Arrange
            const string TEST_TEXT = "ParagraphsAdd";
            var mStream = new MemoryStream();
            PresentationSc presentation = PresentationSc.Open(Resources._001, true);
            TextBoxSc textBox = presentation.Slides[0].Shapes.First(sp => sp.Id == 4).TextBox;
            int originParagraphsCount = textBox.Paragraphs.Count;

            // Act
            ParagraphSc newParagraph = textBox.Paragraphs.Add();
            newParagraph.Text = TEST_TEXT;

            // Assert
            textBox.Paragraphs.Last().Text.Should().BeEquivalentTo(TEST_TEXT);
            textBox.Paragraphs.Should().HaveCountGreaterThan(originParagraphsCount);

            presentation.SaveAs(mStream);
            presentation.Close();
            presentation = PresentationSc.Open(mStream, false);
            textBox = presentation.Slides[0].Shapes.First(sp => sp.Id == 4).TextBox;
            textBox.Paragraphs.Last().Text.Should().BeEquivalentTo(TEST_TEXT);
            textBox.Paragraphs.Should().HaveCountGreaterThan(originParagraphsCount);
        }
    }
}
