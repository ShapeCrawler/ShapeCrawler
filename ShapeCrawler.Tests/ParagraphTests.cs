#if DEBUG

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tests.Helpers;
using ShapeCrawler.Tests.Helpers.Attributes;
using ShapeCrawler.Tests.Properties;
using Xunit;

// ReSharper disable All
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests
{
    public class ParagraphTests : ShapeCrawlerTest, IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public ParagraphTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Theory]
        [SlideParagraphData("autoshape-case003.pptx", 1, "AutoShape 5", 1, 1)]
        [SlideParagraphData("autoshape-case003.pptx", 1, "AutoShape 5", 2, 2)]
        public void Paragraph_IndentLevel_returns_indent_level(IParagraph para, int expectedLevel)
        {
            // Act
            var indentLevel = para.IndentLevel;

            // Assert
            indentLevel.Should().Be(expectedLevel);
        }
        
        [Fact]
        public void Paragraph_Bullet_FontName_Getter_returns_font_name()
        {
            // Arrange
            var shapes = _fixture.Pre002.Slides[1].Shapes;
            var shape3Pr1Bullet = ((IAutoShape)shapes.First(x => x.Id == 3)).TextFrame.Paragraphs[0].Bullet;
            var shape4Pr2Bullet = ((IAutoShape)shapes.First(x => x.Id == 4)).TextFrame.Paragraphs[1].Bullet;

            // Act
            var shape3BulletFontName = shape3Pr1Bullet.FontName;
            var shape4BulletFontName = shape4Pr2Bullet.FontName;

            // Assert
            shape3BulletFontName.Should().BeNull();
            shape4BulletFontName.Should().Be("Calibri");
        }

        [Fact]
        public void Paragraph_Bullet_Type_Getter_returns_bullet_type()
        {
            // Arrange
            var shapeList = _fixture.Pre002.Slides[1].Shapes;
            var shape4 = shapeList.First(x => x.Id == 4);
            var shape5 = shapeList.First(x => x.Id == 5);
            var shape4Pr2Bullet = ((IAutoShape)shape4).TextFrame.Paragraphs[1].Bullet;
            var shape5Pr1Bullet = ((IAutoShape)shape5).TextFrame.Paragraphs[0].Bullet;
            var shape5Pr2Bullet = ((IAutoShape)shape5).TextFrame.Paragraphs[1].Bullet;

            // Act
            var shape5Pr1BulletType = shape5Pr1Bullet.Type;
            var shape5Pr2BulletType = shape5Pr2Bullet.Type;
            var shape4Pr2BulletType = shape4Pr2Bullet.Type;

            // Assert
            shape5Pr1BulletType.Should().Be(SCBulletType.Numbered);
            shape5Pr2BulletType.Should().Be(SCBulletType.Picture);
            shape4Pr2BulletType.Should().Be(SCBulletType.Character);
        }

        [Theory]
        [MemberData(nameof(TestCasesAlignmentGetter))]
        public void Paragraph_Alignment_Getter_returns_text_aligment(IAutoShape autoShape,
            SCTextAlignment expectedAlignment)
        {
            // Arrange
            var paragraph = autoShape.TextFrame.Paragraphs[0];

            // Act
            var textAligment = paragraph.Alignment;

            // Assert
            textAligment.Should().Be(expectedAlignment);
        }

        public static IEnumerable<object[]> TestCasesAlignmentGetter()
        {
            var pptxStream1 = GetTestStream("001.pptx");
            var pres1 = SCPresentation.Open(pptxStream1);
            var autoShape1 = pres1.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 3");
            yield return new object[] { autoShape1, SCTextAlignment.Center };

            var pptxStream2 = GetTestStream("001.pptx");
            var pres2 = SCPresentation.Open(pptxStream2);
            var autoShape2 = pres2.Slides[0].Shapes.GetByName<IAutoShape>("Head 1");
            yield return new object[] { autoShape2, SCTextAlignment.Center };
        }

        [Theory]
        [MemberData(nameof(TestCasesParagraphsAlignmentSetter))]
        public void Paragraph_Alignment_Setter_updates_text_aligment(TestCase testCase)
        {
            // Arrange
            var pres = testCase.Presentation;
            var paragraph = testCase.AutoShape.TextFrame.Paragraphs[0];
            var mStream = new MemoryStream();

            // Act
            paragraph.Alignment = SCTextAlignment.Right;

            // Assert
            paragraph.Alignment.Should().Be(SCTextAlignment.Right);

            pres.SaveAs(mStream);
            testCase.SetPresentation(mStream);
            paragraph = testCase.AutoShape.TextFrame.Paragraphs[0];
            paragraph.Alignment.Should().Be(SCTextAlignment.Right);
        }

        public static IEnumerable<object[]> TestCasesParagraphsAlignmentSetter
        {
            get
            {
                var testCase1 = new TestCase("#1");
                testCase1.PresentationName = "001.pptx";
                testCase1.SlideNumber = 1;
                testCase1.ShapeName = "TextBox 4";
                yield return new[] { testCase1 };

                var testCase2 = new TestCase("#2");
                testCase2.PresentationName = "001.pptx";
                testCase2.SlideNumber = 1;
                testCase2.ShapeName = "Head 1";
                yield return new[] { testCase2 };
            }
        }

        [Fact]
        public void Paragraph_Bullet_Type_Getter_returns_None_value_When_paragraph_doesnt_have_bullet()
        {
            // Arrange
            IAutoShape autoShape = GetAutoShape(presentation: "001.pptx", slideNumber: 1, shapeId: 2);
            var bullet = autoShape.TextFrame.Paragraphs[0].Bullet;

            // Act
            var bulletType = bullet.Type;

            // Assert
            bulletType.Should().Be(SCBulletType.None);
        }

        [Fact]
        public void Paragraph_BulletColorHexAndCharAndSizeProperties_ReturnCorrectValues()
        {
            // Arrange
            var shapeList = _fixture.Pre002.Slides[1].Shapes;
            var shape4 = shapeList.First(x => x.Id == 4);
            var shape4Pr2Bullet = ((IAutoShape)shape4).TextFrame.Paragraphs[1].Bullet;

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
        public void Paragraph_Text_Setter_updates_paragraph_text_and_resize_shape()
        {
            // Arrange
            var pptxStream = GetTestStream("autoshape-case003.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var shape = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 4");
            var paragraph = shape.TextFrame.Paragraphs[0];
            
            // Act
            paragraph.Text = "AutoShape 4 some text";

            // Assert
            shape.Height.Should().Be(46);
            shape.Y.Should().Be(148);
        }

        [Theory]
        [MemberData(nameof(TestCasesParagraphText))]
        public void Paragraph_Text_Setter_updates_paragraph_text(TestElementQuery paragraphQuery, string newText,
            int expectedPortionsCount)
        {
            // Arrange
            var paragraph = paragraphQuery.GetParagraph();
            var mStream = new MemoryStream();
            var pres = paragraphQuery.Presentation;

            // Act
            paragraph.Text = newText;

            // Assert
            paragraph.Text.Should().BeEquivalentTo(newText);
            paragraph.Portions.Should().HaveCount(expectedPortionsCount);

            pres.SaveAs(mStream);
            pres.Close();
            paragraphQuery.Presentation = SCPresentation.Open(mStream);
            paragraph = paragraphQuery.GetParagraph();
            paragraph.Text.Should().BeEquivalentTo(newText);
            paragraph.Portions.Should().HaveCount(expectedPortionsCount);
        }

        public static IEnumerable<object[]> TestCasesParagraphText()
        {
            var paragraphQuery = new TestElementQuery
            {
                SlideIndex = 1,
                ShapeId = 4,
                ParagraphIndex = 2
            };
            paragraphQuery.Presentation = SCPresentation.Open(Resources._002);
            yield return new object[] { paragraphQuery, "Text", 1 };

            paragraphQuery = new TestElementQuery
            {
                SlideIndex = 1,
                ShapeId = 4,
                ParagraphIndex = 2
            };
            paragraphQuery.Presentation = SCPresentation.Open(Resources._002);
            yield return new object[] { paragraphQuery, $"Text{Environment.NewLine}", 1 };

            paragraphQuery = new TestElementQuery
            {
                SlideIndex = 1,
                ShapeId = 4,
                ParagraphIndex = 2
            };
            paragraphQuery.Presentation = SCPresentation.Open(Resources._002);
            yield return new object[] { paragraphQuery, $"Text{Environment.NewLine}Text2", 2 };

            paragraphQuery = new TestElementQuery
            {
                SlideIndex = 1,
                ShapeId = 4,
                ParagraphIndex = 2
            };
            paragraphQuery.Presentation = SCPresentation.Open(Resources._002);
            yield return new object[] { paragraphQuery, $"Text{Environment.NewLine}Text2{Environment.NewLine}", 2 };
        }

        [Fact]
        public void Paragraph_Text_Getter_returns_paragraph_text()
        {
            // Arrange
            ITextFrame textBox1 = ((IAutoShape)_fixture.Pre008.Slides[0].Shapes.First(sp => sp.Id == 37)).TextFrame;
            ITextFrame textBox2 = ((ITable)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0]
                .TextFrame;
            ITextFrame textBox3 = ((ITable)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0]
                .TextFrame;

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
        public void Paragraph_ReplaceText_finds_and_replaces_text()
        {
            // Arrange
            var pptxStream = GetTestStream("autoshape-case003.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var paragraph = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 3").TextFrame!.Paragraphs[0];
            
            // Act
            paragraph.ReplaceText("Some text", "Some text2");

            // Assert
            paragraph.Text.Should().BeEquivalentTo("Some text2");
            var errors = PptxValidator.Validate(pres);
            errors.Should().BeEmpty();
        }

        [Fact]
        public void Paragraph_Portions_counter_returns_number_of_text_portions_in_the_paragraph()
        {
            // Arrange
            var textFrame = _fixture.Pre009.Slides[2].Shapes.GetById<IAutoShape>(2).TextFrame;

            // Act
            var portions = textFrame.Paragraphs[0].Portions;

            // Assert
            portions.Should().HaveCount(2);
        }

        [Theory]
        [SlideShapeData("autoshape-case015.pptx", 1, "TextBox 5", 1.0)]
        [SlideShapeData("autoshape-case015.pptx", 1, "TextBox 4", 1.5)]
        [SlideShapeData("autoshape-case015.pptx", 1, "TextBox 3", 2.0)]
        public void Paragraph_Spacing_LineSpacingLines_returns_line_spacing_in_Lines(IShape shape, double expectedLines)
        {
            // Arrange
            var autoShape = (IAutoShape)shape;
            var paragraph = autoShape.TextFrame!.Paragraphs[0];
            
            // Act
            var spacingLines = paragraph.Spacing.LineSpacingLines;
            
            // Assert
            spacingLines.Should().Be(expectedLines);
            paragraph.Spacing.LineSpacingPoints.Should().BeNull();
        }
        
        [Theory]
        [SlideShapeData("autoshape-case015.pptx", 1, "TextBox 6", 21.6)]
        public void Paragraph_Spacing_LineSpacingPoints_returns_line_spacing_in_Points(IShape shape, double expectedPoints)
        {
            // Arrange
            var autoShape = (IAutoShape)shape;
            var paragraph = autoShape.TextFrame!.Paragraphs[0];
            
            // Act
            var spacingPoints = paragraph.Spacing.LineSpacingPoints;
            
            // Assert
            spacingPoints.Should().Be(expectedPoints);
            paragraph.Spacing.LineSpacingLines.Should().BeNull();
        }
    }
}

#endif