#if DEBUG

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Statics;
using ShapeCrawler.Tests.Helpers;
using ShapeCrawler.Tests.Properties;
using Xunit;

// ReSharper disable All
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests
{
    public class TextBoxTests : ShapeCrawlerTest, IClassFixture<PresentationFixture>
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
            ITextBox textBoxCase2 = ((IAutoShape)_fixture.Pre001.Slides[0].Shapes.First(sp => sp.Id == 5)).TextBox;
            ITextBox textBoxCase3 = ((IAutoShape)_fixture.Pre001.Slides[0].Shapes.First(sp => sp.Id == 6)).TextBox;
            ITextBox textBoxCase5 = ((IAutoShape)_fixture.Pre019.Slides[0].Shapes.First(sp => sp.Id == 2)).TextBox;
            ITextBox textBoxCase6 = ((IAutoShape)_fixture.Pre014.Slides[0].Shapes.First(sp => sp.Id == 61)).TextBox;
            ITextBox textBoxCase7 = ((IAutoShape)_fixture.Pre014.Slides[1].Shapes.First(sp => sp.Id == 5)).TextBox;
            ITextBox textBoxCase8 = ((IAutoShape)_fixture.Pre011.Slides[0].Shapes.First(sp => sp.Id == 54275)).TextBox;
            ITextBox textBoxCase9 = ((IAutoShape)_fixture.Pre008.Slides[0].Shapes.First(sp => sp.Id == 3)).TextBox;
            ITextBox textBoxCase10 = ((IAutoShape)_fixture.Pre021.Slides[3].Shapes.First(sp => sp.Id == 2)).TextBox;
            ITextBox textBoxCase11 = ((IAutoShape)_fixture.Pre012.Slides[0].Shapes.First(sp => sp.Id == 2)).TextBox;
            ITextBox textBoxCase12 = ((IAutoShape)_fixture.Pre012.Slides[0].Shapes.First(sp => sp.Id == 3)).TextBox;
            ITextBox textBoxCase13 = ((IAutoShape)_fixture.Pre011.Slides[0].Shapes.First(sp => sp.Id == 2)).TextBox;
            ITextBox textBoxCase14 = ((ITable)_fixture.Pre001.Slides[1].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0]
                .TextBox;
            ITextBox textBoxCase4 = ((ITable)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0]
                .TextBox;

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
            textBoxCase9.Text.Should().NotBeEmpty();
            textBoxCase10.Text.Should().BeEquivalentTo("test footer");
            textBoxCase11.Text.Should().BeEquivalentTo("Test title text");
            textBoxCase12.Text.Should().BeEquivalentTo("P1 P2");
            textBoxCase13.Text.Should().BeEquivalentTo($"P1{Environment.NewLine}");
            textBoxCase14.Text.Should().BeEquivalentTo("id3");
        }

        [Theory]
        [MemberData(nameof(TestCasesTextSetter))]
        public void Text_Setter_updates_content(TestElementQuery testTextBoxQuery)
        {
            // Arrange
            var pres = testTextBoxQuery.Presentation;
            var textBox = testTextBoxQuery.GetAutoShape().TextBox;
            const string newText = "Test";
            var mStream = new MemoryStream();

            // Act
            textBox.Text = newText;

            // Assert
            textBox.Text.Should().BeEquivalentTo(newText);
            textBox.Paragraphs.Should().HaveCount(1);

            pres.SaveAs(mStream);
            pres.Close();

            testTextBoxQuery.Presentation = SCPresentation.Open(mStream, false);
            textBox = testTextBoxQuery.GetAutoShape().TextBox;
            textBox.Text.Should().BeEquivalentTo(newText);
            textBox.Paragraphs.Should().HaveCount(1);
        }

        public static TheoryData<TestElementQuery> TestCasesTextSetter
        {
            get
            {
                var testCases = new TheoryData<TestElementQuery>();

                var case1 = new TestElementQuery
                {
                    Presentation = SCPresentation.Open(GetTestFileStream("001.pptx"), true),
                    SlideIndex = 0,
                    ShapeId = 3
                };
                testCases.Add(case1);

                var case2 = new TestElementQuery
                {
                    Presentation = SCPresentation.Open(GetTestFileStream("020.pptx"), true),
                    SlideIndex = 2,
                    ShapeId = 8
                };
                testCases.Add(case2);

                var case3 = new TestElementQuery
                {
                    Presentation = SCPresentation.Open(GetTestFileStream("001.pptx"), true),
                    SlideIndex = 1,
                    ShapeId = 2
                };
                testCases.Add(case3);

                return testCases;
            }
        }

        [Fact]
        public void Text_Setter_updates_text_box_content_and_Reduces_font_size_When_text_is_Overflow()
        {
            // Arrange
            var autoShape = GetAutoShape("001.pptx", 1, 9);
            var textBox = autoShape.TextBox;
            var fontSizeBefore = textBox.Paragraphs[0].Portions[0].Font.Size;
            var newText = "Shrink text on overflow";
            PixelConverter.SetDpi(96);

            // Act
            textBox.Text = newText;

            // Assert
            textBox.Text.Should().BeEquivalentTo(newText);
            textBox.Paragraphs[0].Portions[0].Font.Size.Should().Be(8);
        }

        [Fact]
        public void AutofitType_Getter_returns_text_autofit_type()
        {
            // Arrange
            IAutoShape autoShape = GetAutoShape(presentation: "001.pptx", slideNumber: 1, shapeId: 9);
            var textBox = autoShape.TextBox;

            // Act
            var autofitType = textBox.AutofitType;

            // Assert
            autofitType.Should().Be(AutofitType.Shrink);
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
        public void Paragraph_Bullet_Type_Getter_returns_bullet_type()
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
            shape5Pr1BulletType.Should().Be(BulletType.Numbered);
            shape5Pr2BulletType.Should().Be(BulletType.Picture);
            shape4Pr2BulletType.Should().Be(BulletType.Character);
        }

        [Theory]
        [MemberData(nameof(TestCasesAlignmentGetter))]
        public void Paragraph_Alignment_Getter_returns_text_aligment(IAutoShape autoShape,
            TextAlignment expectedAlignment)
        {
            // Arrange
            var paragraph = autoShape.TextBox.Paragraphs[0];

            // Act
            var textAligment = paragraph.Alignment;

            // Assert
            textAligment.Should().Be(expectedAlignment);
        }

        public static IEnumerable<object[]> TestCasesAlignmentGetter()
        {
            var pptxStream = GetTestFileStream("001.pptx");
            var presentation = SCPresentation.Open(pptxStream, false);
            var autoShape = presentation.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 3");
            yield return new object[] { autoShape, TextAlignment.Center };

            pptxStream = GetTestFileStream("001.pptx");
            presentation = SCPresentation.Open(pptxStream, false);
            autoShape = presentation.Slides[0].Shapes.GetByName<IAutoShape>("Head 1");
            yield return new object[] { autoShape, TextAlignment.Center };
        }

        [Fact]
        public void Paragraph_Alignment_Setter_updates_text_aligment()
        {
            // Arrange
            var pptxStream = GetTestFileStream("001.pptx");
            var originPresentation = SCPresentation.Open(pptxStream, true);
            var autoShape = originPresentation.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 4");
            var paragraph = autoShape.TextBox.Paragraphs[0];

            // Act
            paragraph.Alignment = TextAlignment.Right;

            // Assert
            paragraph.Alignment.Should().Be(TextAlignment.Right);

            var modifiedPresentation = SaveAndOpenPresentation(originPresentation);
            autoShape = originPresentation.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 4");
            paragraph = autoShape.TextBox.Paragraphs[0];
            paragraph.Alignment.Should().Be(TextAlignment.Right);
        }

        [Fact]
        public void Paragraph_Bullet_Type_Getter_returns_None_value_When_paragraph_doesnt_have_bullet()
        {
            // Arrange
            IAutoShape autoShape = GetAutoShape(presentation: "001.pptx", slideNumber: 1, shapeId: 2);
            var bullet = autoShape.TextBox.Paragraphs[0].Bullet;

            // Act
            var bulletType = bullet.Type;

            // Assert
            bulletType.Should().Be(BulletType.None);
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
        public void Paragraph_Text_Setter_ThrowsException_When_paragraph_was_removed()
        {
            IPresentation presentation = SCPresentation.Open(Properties.Resources._020, true);
            IAutoShape autoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 8);
            ITextBox textBox = autoShape.TextBox;
            IParagraph paragraph = textBox.Paragraphs.First();
            textBox.Text = "new box content";

            // Act-Assert
            paragraph.Invoking(p => p.Text = "new paragraph text")
                .Should().Throw<ElementIsRemovedException>(
                    "because paragraph was being removed while changing box content.");
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
            paragraphQuery.Presentation = SCPresentation.Open(mStream, false);
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
            paragraphQuery.Presentation = SCPresentation.Open(Resources._002, true);
            yield return new object[] { paragraphQuery, "Text", 1 };

            paragraphQuery = new TestElementQuery
            {
                SlideIndex = 1,
                ShapeId = 4,
                ParagraphIndex = 2
            };
            paragraphQuery.Presentation = SCPresentation.Open(Resources._002, true);
            yield return new object[] { paragraphQuery, $"Text{Environment.NewLine}", 1 };

            paragraphQuery = new TestElementQuery
            {
                SlideIndex = 1,
                ShapeId = 4,
                ParagraphIndex = 2
            };
            paragraphQuery.Presentation = SCPresentation.Open(Resources._002, true);
            yield return new object[] { paragraphQuery, $"Text{Environment.NewLine}Text2", 2 };

            paragraphQuery = new TestElementQuery
            {
                SlideIndex = 1,
                ShapeId = 4,
                ParagraphIndex = 2
            };
            paragraphQuery.Presentation = SCPresentation.Open(Resources._002, true);
            yield return new object[] { paragraphQuery, $"Text{Environment.NewLine}Text2{Environment.NewLine}", 2 };
        }

        [Fact]
        public void ParagraphText_GetterReturnsParagraphText()
        {
            // Arrange
            ITextBox textBox1 = ((IAutoShape)_fixture.Pre008.Slides[0].Shapes.First(sp => sp.Id == 37)).TextBox;
            ITextBox textBox2 = ((ITable)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0]
                .TextBox;
            ITextBox textBox3 = ((ITable)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0]
                .TextBox;

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
        public void ParagraphPortions_CollectionCounterReturnsNumberOfTextPortionsInTheParagraph()
        {
            // Arrange
            ITextBox textBox = ((IAutoShape)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 2)).TextBox;

            // Act
            IEnumerable<IPortion> paragraphPortions = textBox.Paragraphs[0].Portions;

            // Assert
            paragraphPortions.Should().HaveCount(2);
        }

        [Theory]
        [MemberData(nameof(TestCasesParagraphsCount))]
        public void Paragraphs_Count_returns_number_of_paragraphs_in_the_text_box(TestCase<ITextBox, int> testCase)
        {
            // Arrange
            var textBox = testCase.Param1;
            var expectedParaCount = testCase.Param2;
            var paragraphs = textBox.Paragraphs;

            // Act
            var actualParaCount = paragraphs.Count;

            // Assert
            actualParaCount.Should().Be(expectedParaCount);
        }

        public static IEnumerable<object[]> TestCasesParagraphsCount
        {
            get
            {
                var pptxStream1 = GetTestFileStream("009_table.pptx");
                var pres1 = SCPresentation.Open(pptxStream1, false);
                var autoShape1 = pres1.Slides[2].Shapes.GetById<IAutoShape>(2);
                var textBox1 = autoShape1.TextBox;
                var testCase1 = new TestCase<ITextBox, int>(1, textBox1, 1);
                yield return new object[] { testCase1 };

                var pptxStream2 = GetTestFileStream("020.pptx");
                var pres2 = SCPresentation.Open(pptxStream2, false);
                var autoShape2 = pres2.Slides[2].Shapes.GetById<IAutoShape>(8);
                var textBox2 = autoShape2.TextBox;
                var testCase2 = new TestCase<ITextBox, int>(2, textBox2, 1);
                yield return new object[] { testCase2 };

                var pptxStream3 = GetTestFileStream("009_table.pptx");
                var pres3 = SCPresentation.Open(pptxStream3, false);
                var table3 = pres3.Slides[2].Shapes.GetById<ITable>(3);
                var textBox3 = table3.Rows[0].Cells[0].TextBox;
                var testCase3 = new TestCase<ITextBox, int>(3, textBox3, 2);
                yield return new object[] { testCase3 };
                
                var pptxStream4 = GetTestFileStream("001.pptx");
                var pres4 = SCPresentation.Open(pptxStream4, false);
                var autoShape4 = pres4.Slides[1].Shapes.GetById<IAutoShape>(2);
                var textBox4 = autoShape4.TextBox;
                var testCase4 = new TestCase<ITextBox, int>(4, textBox4, 0);
                yield return new object[] { testCase4 };
            }
        }

        [Fact]
        public void ParagraphsAdd_AddsANewTextParagraphAtTheEndOfTheTextBoxAndReturnsAddedParagraph()
        {
            // Arrange
            const string TEST_TEXT = "ParagraphsAdd";
            var mStream = new MemoryStream();
            IPresentation presentation = SCPresentation.Open(Resources._001, true);
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
        public void
            Paragraphs_Add_returns_a_new_added_paragraph_When_paragraph_has_been_added_after_text_box_content_changed()
        {
            var pres = SCPresentation.Open(Properties.Resources._001, true);
            var autoShape = (IAutoShape)pres.Slides[0].Shapes.First(sp => sp.Id == 3);
            var textBox = autoShape.TextBox;
            var paragraphs = textBox.Paragraphs;
            var paragraph = textBox.Paragraphs.First();

            // Act
            textBox.Text = "A new text";
            var newParagraph = paragraphs.Add();

            // Assert
            newParagraph.Should().NotBeNull();
        }
    }
}

#endif