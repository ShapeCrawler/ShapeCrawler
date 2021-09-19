using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tables;
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
            ITextBox textBoxCase1 = ((IAutoShape)_fixture.Pre009.Slides[3].Shapes.First(sp => sp.Id == 2)).TextBox;
            ITextBox textBoxCase2 = ((IAutoShape)_fixture.Presentation.Slides[0].Shapes.First(sp => sp.Id == 5)).TextBox;
            ITextBox textBoxCase3 = ((IAutoShape)_fixture.Presentation.Slides[0].Shapes.First(sp => sp.Id == 6)).TextBox;
            ITextBox textBoxCase5 = ((IAutoShape)_fixture.Pre019.Slides[0].Shapes.First(sp => sp.Id == 2)).TextBox;
            ITextBox textBoxCase6 = ((IAutoShape)_fixture.Pre014.Slides[0].Shapes.First(sp => sp.Id == 61)).TextBox;
            ITextBox textBoxCase7 = ((IAutoShape)_fixture.Pre014.Slides[1].Shapes.First(sp => sp.Id == 5)).TextBox;
            ITextBox textBoxCase8 = ((IAutoShape)_fixture.Pre011.Slides[0].Shapes.First(sp => sp.Id == 54275)).TextBox;
            ITextBox textBoxCase9 = ((IAutoShape)_fixture.Pre008.Slides[0].Shapes.First(sp => sp.Id == 3)).TextBox;
            ITextBox textBoxCase10 = ((IAutoShape)_fixture.Pre021.Slides[3].Shapes.First(sp => sp.Id == 2)).TextBox;
            ITextBox textBoxCase11 = ((IAutoShape)_fixture.Pre012.Slides[0].Shapes.First(sp => sp.Id == 2)).TextBox;
            ITextBox textBoxCase12 = ((IAutoShape)_fixture.Pre012.Slides[0].Shapes.First(sp => sp.Id == 3)).TextBox;
            ITextBox textBoxCase13 = ((IAutoShape)_fixture.Pre011.Slides[0].Shapes.First(sp => sp.Id == 2)).TextBox;
            ITextBox textBoxCase14 = ((ITable)_fixture.Presentation.Slides[1].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0].TextBox;
            ITextBox textBoxCase4 = ((ITable)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0].TextBox;

            // Act-Assert
            textBoxCase1.Text.Should().BeEquivalentTo("Title text");
            textBoxCase2.Text.Should().BeEquivalentTo(" id5-Text1");
            textBoxCase3.Text.Should().BeEquivalentTo($"id6-Text1{Environment.NewLine}Text2");
            textBoxCase4.Text.Should().BeEquivalentTo($"0:0_p1_lvl1{Environment.NewLine}0:0_p2_lvl2");
            textBoxCase5.Text.Should().BeEquivalentTo("1");
            textBoxCase6.Text.Should().BeEquivalentTo($"test1{Environment.NewLine}test2{Environment.NewLine}" +
                                                   $"test3{Environment.NewLine}test4{Environment.NewLine}test5");
            textBoxCase7.Text.Should().BeEquivalentTo("Test subtitle");
            textBoxCase8.Text.Should().BeEquivalentTo("Jan 2018");
            textBoxCase9.Text.Should().BeEquivalentTo("25.01.2020");
            textBoxCase10.Text.Should().BeEquivalentTo("test footer");
            textBoxCase11.Text.Should().BeEquivalentTo("Test title text");
            textBoxCase12.Text.Should().BeEquivalentTo("P1 P2");
            textBoxCase13.Text.Should().BeEquivalentTo($"P1{Environment.NewLine}");
            textBoxCase14.Text.Should().BeEquivalentTo("id3");
        }

        [Fact]
        public void Text_SetterChangesTextBoxContent()
        {
            // Arrange
            IPresentation presentation = SCPresentation.Open(Resources._001, true);
            ITextBox textBox = ((IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 3)).TextBox;
            const string newText = "Test";
            var mStream = new MemoryStream();

            // Act
            textBox.Text = newText;

            // Assert
            textBox.Text.Should().BeEquivalentTo(newText);
            textBox.Paragraphs.Should().HaveCount(1);
            
            presentation.SaveAs(mStream);
            presentation.Close();

            presentation = SCPresentation.Open(mStream, false);
            textBox = ((IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 3)).TextBox;
            textBox.Text.Should().BeEquivalentTo(newText);
            textBox.Paragraphs.Should().HaveCount(1);
        }

        [Fact]
        public void Text_SetterChangesTextBoxContent_WhenTheFirstParagraphIsEmpty()
        {
            // Arrange
            IPresentation presentation = SCPresentation.Open(Resources._020, true);
            ITextBox textBox = ((IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 8)).TextBox;
            const string newText = "Test";
            var mStream = new MemoryStream();

            // Act
            textBox.Text = newText;

            // Assert
            textBox.Text.Should().BeEquivalentTo(newText);
            textBox.Paragraphs.Should().HaveCount(1);

            presentation.SaveAs(mStream);
            presentation.Close();
            presentation = SCPresentation.Open(mStream, false);
            textBox = ((IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 8)).TextBox;
            textBox.Text.Should().BeEquivalentTo(newText);
            textBox.Paragraphs.Should().HaveCount(1);
        }

        [Fact]
        public void Shape_IsAutoShape()
        {
            // Arrange
            IShape shapeCase1 = _fixture.Pre008.Slides[0].Shapes.First(sp => sp.Id == 3);
            IShape shapeCase2 = _fixture.Pre021.Slides[3].Shapes.First(sp => sp.Id == 2);
            IShape shapeCase3 = _fixture.Pre011.Slides[0].Shapes.First(sp => sp.Id == 54275);

            // Act
            var autoShapeCase1 = shapeCase1 as IAutoShape;
            var autoShapeCase2 = shapeCase2 as IAutoShape;
            var autoShapeCase3 = shapeCase3 as IAutoShape;

            // Assert
            autoShapeCase1.Should().NotBeNull();
            autoShapeCase2.Should().NotBeNull();
            autoShapeCase3.Should().NotBeNull();
        }

        [Fact]
        public void ParagraphBulletFontNameProperty_ReturnsFontName()
        {
            // Arrange
            var shapes = _fixture.Pre002.Slides[1].Shapes;
            var shape3Pr1Bullet = ((IAutoShape)shapes.First(x => x.Id == 3)).TextBox.Paragraphs[0].Bullet;
            var shape4Pr2Bullet = ((IAutoShape)shapes.First(x => x.Id == 4)).TextBox.Paragraphs[1].Bullet;

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
            var shape4Pr2Bullet = ((IAutoShape)shape4).TextBox.Paragraphs[1].Bullet;
            var shape5Pr1Bullet = ((IAutoShape)shape5).TextBox.Paragraphs[0].Bullet;
            var shape5Pr2Bullet = ((IAutoShape)shape5).TextBox.Paragraphs[1].Bullet;

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
            var shape4Pr2Bullet = ((IAutoShape)shape4).TextBox.Paragraphs[1].Bullet;

            // Act
            var bulletColorHex = shape4Pr2Bullet.ColorHex;
            var bulletChar = shape4Pr2Bullet.Character;
            var bulletSize = shape4Pr2Bullet.Size;

            // Assert
            bulletColorHex.Should().Be("C00000");
            bulletChar.Should().Be("'");
            bulletSize.Should().Be(120);
        }

        [Fact]
        public void ParagraphTextSetter_ThrowsException_WhenParagraphWasRemoved()
        {
            IPresentation presentation = SCPresentation.Open(Properties.Resources._020, true);
            IAutoShape autoShape = (IAutoShape) presentation.Slides[2].Shapes.First(sp => sp.Id == 8);
            ITextBox textBox = autoShape.TextBox;
            IParagraph paragraph = textBox.Paragraphs.First();
            textBox.Text = "new box content";

            // Act-Assert
            paragraph.Invoking(p => p.Text = "new paragraph text")
                .Should().Throw<ElementIsRemovedException>("because paragraph was being removed while changing box content.");
        }

        [Theory]
        [MemberData(nameof(TestCasesParagraphText))]
        public void ParagraphText_SetterChangesParagraphText(
            SCPresentation presentation, 
            SlideElementQuery prRequest, 
            string newPrText,
            int expectedNumPortions)
        {
            // Arrange
            IParagraph paragraph = TestHelper.GetParagraph(presentation, prRequest);
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
            var paragraphRequest = new SlideElementQuery
            {
                SlideIndex = 1,
                ShapeId = 4,
                ParagraphIndex = 1
            };
            SCPresentation presentation;
            paragraphRequest.ParagraphIndex = 2;

            presentation = SCPresentation.Open(Resources._002, true);
            yield return new object[] { presentation, paragraphRequest, "Text", 1};

            presentation = SCPresentation.Open(Resources._002, true);
            yield return new object[] { presentation, paragraphRequest, $"Text{Environment.NewLine}", 1};

            presentation = SCPresentation.Open(Resources._002, true);
            yield return new object[] { presentation, paragraphRequest, $"Text{Environment.NewLine}Text2", 2};

            presentation = SCPresentation.Open(Resources._002, true);
            yield return new object[] { presentation, paragraphRequest, $"Text{Environment.NewLine}Text2{Environment.NewLine}", 2 };
        }

        [Fact]
        public void ParagraphText_GetterReturnsParagraphText()
        {
            // Arrange
            ITextBox textBox1 = ((IAutoShape)_fixture.Pre008.Slides[0].Shapes.First(sp => sp.Id == 37)).TextBox;
            ITextBox textBox2 = ((ITable)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0].TextBox;
            ITextBox textBox3 = ((ITable)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0].TextBox;

            // Act
            string paragraphTextCase1 = textBox1.Paragraphs[0].Text;
            string paragraphTextCase2 = textBox1.Paragraphs[1].Text;
            string paragraphTextCase3 = textBox2.Paragraphs[0].Text;

            // Assert
            paragraphTextCase1.Should().BeEquivalentTo("P1t1 P1t2");
            paragraphTextCase2.Should().BeEquivalentTo("p2");
            paragraphTextCase3.Should().BeEquivalentTo("0:0_p1_lvl1");
        }

        [Fact]
        public void Paragraphs_CollectionCounterReturnsNumberOfParagraphsInTheTextFrame()
        {
            // Arrange
            ITextBox textBoxCase1 = ((IAutoShape)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 2)).TextBox;
            ITextBox textBoxCase2 = ((IAutoShape)_fixture.Pre020.Slides[2].Shapes.First(sp => sp.Id == 8)).TextBox;

            // Act
            IEnumerable<IParagraph> paragraphsC1 = textBoxCase1.Paragraphs;
            IEnumerable<IParagraph> paragraphsC2 = textBoxCase2.Paragraphs;

            // Assert
            paragraphsC1.Should().HaveCount(1);
            paragraphsC2.Should().HaveCount(2);
        }

        [Fact]
        public void ParagraphPortions_CollectionCounterReturnsNumberOfTextPortionsInTheParagraph()
        {
            // Arrange
            ITextBox textBox = ((IAutoShape)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 2)).TextBox;

            // Act
            IEnumerable<IPortion> paragraphPortions = textBox.Paragraphs[0].Portions;

            // Assert
            paragraphPortions.Should().HaveCount(2);
        }

        [Fact]
        public void ParagraphsCount_ReturnsTwo_WhenNumberOfParagraphsInCellTextBoxIsTwo()
        {
            // Arrange
            ITable table = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3) as ITable;
            ITextBox textBox = table.Rows[0].Cells[0].TextBox;

            // Act-Assert
            textBox.Paragraphs.Should().HaveCount(2);
        }

        [Fact]
        public void ParagraphsAdd_AddsANewTextParagraphAtTheEndOfTheTextBoxAndReturnsAddedParagraph()
        {
            // Arrange
            const string TEST_TEXT = "ParagraphsAdd";
            var mStream = new MemoryStream();
            SCPresentation presentation = SCPresentation.Open(Resources._001, true);
            ITextBox textBox = ((IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 4)).TextBox;
            int originParagraphsCount = textBox.Paragraphs.Count;

            // Act
            IParagraph newParagraph = textBox.Paragraphs.Add();
            newParagraph.Text = TEST_TEXT;

            // Assert
            textBox.Paragraphs.Last().Text.Should().BeEquivalentTo(TEST_TEXT);
            textBox.Paragraphs.Should().HaveCountGreaterThan(originParagraphsCount);

            presentation.SaveAs(mStream);
            presentation = SCPresentation.Open(mStream, false);
            textBox = ((IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 4)).TextBox;
            textBox.Paragraphs.Last().Text.Should().BeEquivalentTo(TEST_TEXT);
            textBox.Paragraphs.Should().HaveCountGreaterThan(originParagraphsCount);
        }

        [Fact]
        public void ParagraphsAdd_AddsANewTextParagraphAtTheEndOfTheTextBoxAndReturnsAddedParagraph_WhenParagraphIsAddedAfterTextBoxContentChanged()
        {
            IPresentation presentation = SCPresentation.Open(Properties.Resources._001, true);
            IAutoShape autoShape = (IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 3);
            ITextBox textBox = autoShape.TextBox;
            IParagraphCollection paragraphs = textBox.Paragraphs;
            IParagraph paragraph = textBox.Paragraphs.First();
            textBox.Text = "A new text";

            // Act
            IParagraph newParagraph = paragraphs.Add();

            // Assert
            newParagraph.Should().NotBeNull();
        }
    }
}