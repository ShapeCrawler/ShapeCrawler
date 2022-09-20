using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Collections;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Tests.Helpers;
using ShapeCrawler.Tests.Properties;
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
            IPortion portion = ((ITable)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0]
                .TextFrame
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
            IPresentation presentation = SCPresentation.Open(TestFiles.Presentations.pre001);
            IAutoShape autoShape = (IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 5);
            IPortionCollection portions = autoShape.TextFrame.Paragraphs[0].Portions;
            IPortion portion = portions[0];
            portions.Remove(portion);

            // Act-Assert
            portion.Invoking(p => p.Text = "new text").Should().Throw<Exception>();
        }

        [Fact]
        public void Text_Setter_updates_text()
        {
            // Arrange
            var pptxStream = GetTestStream("autoshape-case001.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var autoShape = pres.SlideMasters[0].Shapes.GetByName<IAutoShape>("AutoShape 1");
            var portion = autoShape.TextFrame.Paragraphs[0].Portions[0];

            // Act
            portion.Text = "test";

            // Assert
            portion.Text.Should().Be("test");
        }

        [Theory]
        [MemberData(nameof(TestCasesHyperlinkSetter))]
        public void Hyperlink_Setter_sets_hyperlink(string pptxFile, string shapeName)
        {
            // Arrange
            var pptxStream = GetTestStream(pptxFile);
            var presentation = SCPresentation.Open(pptxStream);
            var autoShape = presentation.Slides[0].Shapes.GetByName<IAutoShape>(shapeName);
            var portion = autoShape.TextFrame.Paragraphs[0].Portions[0];

            // Act
            portion.Hyperlink = "https://github.com/ShapeCrawler/ShapeCrawler";

            // Assert
            presentation.Save();
            presentation.Close();
            presentation = SCPresentation.Open(pptxStream);
            autoShape = presentation.Slides[0].Shapes.GetByName<IAutoShape>(shapeName);
            portion = autoShape.TextFrame.Paragraphs[0].Portions[0];
            portion.Hyperlink.Should().Be("https://github.com/ShapeCrawler/ShapeCrawler");
        }

        public static IEnumerable<object[]> TestCasesHyperlinkSetter()
        {
            yield return new[] { "001.pptx", "TextBox 3" };
            yield return new[] { "autoshape-case001.pptx", "AutoShape 1" };
            yield return new[] { "autoshape-case002.pptx", "AutoShape 1" };
        }

        [Fact]
        public void Hyperlink_Setter_sets_hyperlink_for_two_shape_on_the_Same_slide()
        {
            // Arrange
            var pptxStream = GetTestStream("001.pptx");
            var presentation = SCPresentation.Open(pptxStream);
            var textBox3 = presentation.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 3");
            var textBox4 = presentation.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 4");
            var portion3 = textBox3.TextFrame.Paragraphs[0].Portions[0];
            var portion4 = textBox4.TextFrame.Paragraphs[0].Portions[0];

            // Act
            portion3.Hyperlink = "https://github.com/ShapeCrawler/ShapeCrawler";
            portion4.Hyperlink = "https://github.com/ShapeCrawler/ShapeCrawler";

            // Assert
            portion3.Hyperlink.Should().Be("https://github.com/ShapeCrawler/ShapeCrawler");
            portion4.Hyperlink.Should().Be("https://github.com/ShapeCrawler/ShapeCrawler");
        }

        [Fact]
        public void Bullet_PropertyChangedToCharacter_WhenValueEqualsSetPassed()
        {
            // Arrange
            var mStream = new MemoryStream();
            IPresentation presentation = SCPresentation.Open(Resources._020);
            IAutoShape placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
            IParagraph paragraph = placeholderAutoShape.TextFrame.Paragraphs.Add();

            // Act
            paragraph.Bullet.Type = SCBulletType.Character;
            paragraph.Bullet.Character = "*";
            paragraph.Bullet.Size = 100;
            paragraph.Bullet.FontName = "Tahoma";

            // Assert
            paragraph.Bullet.Type.Should().Be(SCBulletType.Character);
            paragraph.Bullet.Character.Should().Be("*");
            paragraph.Bullet.Size.Should().Be(100);
            paragraph.Bullet.FontName.Should().Be("Tahoma");

            presentation.SaveAs(mStream);

            presentation = SCPresentation.Open(mStream);
            placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
            paragraph = placeholderAutoShape.TextFrame.Paragraphs.Last();
            paragraph.Bullet.Type.Should().Be(SCBulletType.Character);
            paragraph.Bullet.Character.Should().Be("*");
            paragraph.Bullet.Size.Should().Be(100);
            paragraph.Bullet.FontName.Should().Be("Tahoma");
        }

        [Fact]
        public void Bullet_PropertyChangedToNumbered_WhenValueEqualsSetPassed()
        {
            // Arrange
            var mStream = new MemoryStream();
            IPresentation presentation = SCPresentation.Open(Resources._020);
            IAutoShape placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
            IParagraph paragraph = placeholderAutoShape.TextFrame.Paragraphs.Add();

            // Act
            paragraph.Bullet.Type = SCBulletType.Numbered;
            paragraph.Bullet.Size = 100;
            paragraph.Bullet.FontName = "Tahoma";

            // Assert
            paragraph.Bullet.Type.Should().Be(SCBulletType.Numbered);
            paragraph.Bullet.Size.Should().Be(100);
            paragraph.Bullet.FontName.Should().Be("Tahoma");

            presentation.SaveAs(mStream);

            presentation = SCPresentation.Open(mStream);
            placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
            paragraph = placeholderAutoShape.TextFrame.Paragraphs.Last();
            paragraph.Bullet.Type.Should().Be(SCBulletType.Numbered);
            paragraph.Bullet.Size.Should().Be(100);
            paragraph.Bullet.FontName.Should().Be("Tahoma");
        }

    }
}