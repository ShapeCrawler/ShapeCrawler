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
    public class TextFrameTests : ShapeCrawlerTest, IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public TextFrameTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Text_Getter_returns_text_of_table_Cell()
        {
            // Arrange
            var textFrame1 = ((IAutoShape)_fixture.Pre008.Slides[0].Shapes.First(sp => sp.Id == 3)).TextFrame;
            var textFrame2 = ((ITable)_fixture.Pre001.Slides[1].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0]
                .TextFrame;
            var textFrame3 = ((ITable)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0]
                .TextFrame;
            
            // Act
            var text1 = textFrame1.Text;
            var text2 = textFrame2.Text;
            var text3 = textFrame3.Text;

            // Act
            text1.Should().NotBeEmpty();
            text2.Should().BeEquivalentTo("id3");
            text3.Should().BeEquivalentTo($"0:0_p1_lvl1{Environment.NewLine}0:0_p2_lvl2");
        }
        
        [Theory]
        [SlideShapeData("009_table.pptx", 4, 2, "Title text")]
        [SlideShapeData("001.pptx", 1, 5, " id5-Text1")]
        [SlideShapeData("019.pptx", 1, 2, "1")]
        [SlideShapeData("014.pptx", 2, 5, "Test subtitle")]
        [SlideShapeData("011_dt.pptx", 1, 54275, "Jan 2018")]
        [SlideShapeData("021.pptx", 4, 2, "test footer")]
        [SlideShapeData("012_title-placeholder.pptx", 1, 2, "Test title text")]
        [SlideShapeData("012_title-placeholder.pptx", 1, 3, "P1 P2")]
        public void Text_Getter_returns_text(IShape shape, string expectedText)
        {
            // Arrange
            var textFrame = ((IAutoShape)shape).TextFrame;

            // Act
            var text = textFrame.Text;

            // Assert
            text.Should().BeEquivalentTo(expectedText);
        }

        [Theory]
        [MemberData(nameof(TextGetterTestCases))]
        public void Text_Getter_returns_text_with_New_Line(TestCase testCase)
        {
            // Arrange
            var textFrame = testCase.AutoShape.TextFrame;
            var expectedText = testCase.ExpectedString;

            // Act
            var text = textFrame.Text;

            // Assert
            text.Should().BeEquivalentTo(expectedText);
        }
        
        public static IEnumerable<object[]> TextGetterTestCases
        {
            get
            {
                var testCase3 = new TestCase("#3");
                testCase3.PresentationName = "001.pptx";
                testCase3.SlideNumber = 1;
                testCase3.ShapeId = 6;
                testCase3.ExpectedString = $"id6-Text1{Environment.NewLine}Text2";
                yield return new object[] { testCase3 };
                
                var testCase5 = new TestCase("#5");
                testCase5.PresentationName = "014.pptx";
                testCase5.SlideNumber = 1;
                testCase5.ShapeId = 61;
                testCase5.ExpectedString = $"test1{Environment.NewLine}test2{Environment.NewLine}" +
                                           $"test3{Environment.NewLine}test4{Environment.NewLine}test5";
                yield return new object[] { testCase5 };
                
                var testCase11 = new TestCase("#11");
                testCase11.PresentationName = "011_dt.pptx";
                testCase11.SlideNumber = 1;
                testCase11.ShapeId = 2;
                testCase11.ExpectedString = $"P1{Environment.NewLine}";
                yield return new object[] { testCase11 };
            }
        }

        [Fact]
        public void Text_Setter_can_update_content_multiple_times()
        {
            // Arrange
            var pres = SCPresentation.Open(Properties.Resources.autoshape_case005_text_frame);
            var textFrame = pres.Slides.First().Shapes.OfType<IAutoShape>().First().TextFrame;

            // Act
            textFrame.Text = textFrame.Text.Replace("{{replace_this}}", "confirm this");
            textFrame.Text = textFrame.Text.Replace("{{replace_that}}", "confirm that");

            var modifiedPres = new MemoryStream();
            pres.SaveAs(modifiedPres);
            pres.Close();
            pres = SCPresentation.Open(modifiedPres);

            // Assert
            var changedTextFrame = pres.Slides.First().Shapes.OfType<IAutoShape>().First().TextFrame;

            changedTextFrame.Text.Should().ContainAll("confirm this", "confirm that");
        }
        
        [Theory]
        [MemberData(nameof(TestCasesTextSetter))]
        public void Text_Setter_updates_content(TestElementQuery testTextBoxQuery)
        {
            // Arrange
            var pres = testTextBoxQuery.Presentation;
            var textBox = testTextBoxQuery.GetAutoShape().TextFrame;
            const string newText = "Test";
            var mStream = new MemoryStream();

            // Act
            textBox.Text = newText;

            // Assert
            textBox.Text.Should().BeEquivalentTo(newText);
            textBox.Paragraphs.Should().HaveCount(1);

            pres.SaveAs(mStream);
            pres.Close();

            testTextBoxQuery.Presentation = SCPresentation.Open(mStream);
            textBox = testTextBoxQuery.GetAutoShape().TextFrame;
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
                    Presentation = SCPresentation.Open(GetTestStream("001.pptx")),
                    SlideIndex = 0,
                    ShapeId = 3
                };
                testCases.Add(case1);
                
                var case2 = new TestElementQuery
                {
                    Presentation = SCPresentation.Open(GetTestStream("020.pptx")),
                    SlideIndex = 2,
                    ShapeId = 8
                };
                testCases.Add(case2);
                
                var case3 = new TestElementQuery
                {
                    Presentation = SCPresentation.Open(GetTestStream("001.pptx")),
                    SlideNumber = 2,
                    ShapeName = "Header 1",
                };
                testCases.Add(case3);
                
                var case4 = new TestElementQuery
                {
                    Presentation = SCPresentation.Open(GetTestStream("autoshape-case004_subtitle.pptx")),
                    SlideNumber = 1,
                    ShapeName = "Subtitle 1",
                };
                testCases.Add(case4);
                
                var case5 = new TestElementQuery
                {
                    Presentation = SCPresentation.Open(GetTestStream("autoshape-case008_text-frame.pptx")),
                    SlideNumber = 1,
                    ShapeName = "AutoShape 1",
                };
                testCases.Add(case5);

                return testCases;
            }
        }

        [Fact]
        public void Text_Setter_updates_text_box_content_and_Reduces_font_size_When_text_is_Overflow()
        {
            // Arrange
            var pptxStream = GetTestStream("001.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var textBox = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 8");
            var textFrame = textBox.TextFrame;
            var fontSizeBefore = textFrame.Paragraphs[0].Portions[0].Font.Size;
            var newText = "Shrink text on overflow";

            // Act
            textFrame.Text = newText;

            // Assert
            textFrame.Text.Should().BeEquivalentTo(newText);
            textFrame.Paragraphs[0].Portions[0].Font.Size.Should().Be(8);
        }

        [Fact]
        public void AutofitType_Getter_returns_text_autofit_type()
        {
            // Arrange
            IAutoShape autoShape = GetAutoShape(presentation: "001.pptx", slideNumber: 1, shapeId: 9);
            var textBox = autoShape.TextFrame;

            // Act
            var autofitType = textBox.AutoFitType;

            // Assert
            autofitType.Should().Be(SCAutoFitType.Shrink);
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
        public void ParagraphBulletColorHexAndCharAndSizeProperties_ReturnCorrectValues()
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
        public void Paragraph_ReplaceText_finds_and_replases_text()
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
        public void ParagraphPortions_CollectionCounterReturnsNumberOfTextPortionsInTheParagraph()
        {
            // Arrange
            ITextFrame textBox = ((IAutoShape)_fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 2)).TextFrame;

            // Act
            IEnumerable<IPortion> paragraphPortions = textBox.Paragraphs[0].Portions;

            // Assert
            paragraphPortions.Should().HaveCount(2);
        }

        [Theory]
        [MemberData(nameof(TestCasesParagraphsCount))]
        public void Paragraphs_Count_returns_number_of_paragraphs_in_the_text_box(TestCase<ITextFrame, int> testCase)
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
                var pptxStream1 = GetTestStream("009_table.pptx");
                var pres1 = SCPresentation.Open(pptxStream1);
                var autoShape1 = pres1.Slides[2].Shapes.GetById<IAutoShape>(2);
                var textBox1 = autoShape1.TextFrame;
                var testCase1 = new TestCase<ITextFrame, int>(1, textBox1, 1);
                yield return new object[] { testCase1 };

                var pptxStream2 = GetTestStream("020.pptx");
                var pres2 = SCPresentation.Open(pptxStream2);
                var autoShape2 = pres2.Slides[2].Shapes.GetById<IAutoShape>(8);
                var textBox2 = autoShape2.TextFrame;
                var testCase2 = new TestCase<ITextFrame, int>(2, textBox2, 2);
                yield return new object[] { testCase2 };

                var pptxStream3 = GetTestStream("009_table.pptx");
                var pres3 = SCPresentation.Open(pptxStream3);
                var table3 = pres3.Slides[2].Shapes.GetById<ITable>(3);
                var textBox3 = table3.Rows[0].Cells[0].TextFrame;
                var testCase3 = new TestCase<ITextFrame, int>(3, textBox3, 2);
                yield return new object[] { testCase3 };

                var pptxStream4 = GetTestStream("001.pptx");
                var pres4 = SCPresentation.Open(pptxStream4);
                var autoShape4 = pres4.Slides[1].Shapes.GetById<IAutoShape>(2);
                var textBox4 = autoShape4.TextFrame;
                var testCase4 = new TestCase<ITextFrame, int>(4, textBox4, 1);
                yield return new object[] { testCase4 };
            }
        }

        [Fact]
        public void Paragraphs_Add_adds_new_text_paragraph_at_the_end_And_returns_added_paragraph()
        {
            // Arrange
            const string TEST_TEXT = "ParagraphsAdd";
            var mStream = new MemoryStream();
            var pres = SCPresentation.Open(Resources._001);
            var textFrame = ((IAutoShape)pres.Slides[0].Shapes.First(sp => sp.Id == 4)).TextFrame;
            int originParagraphsCount = textFrame.Paragraphs.Count;

            // Act
            var addedPara = textFrame.Paragraphs.Add();
            addedPara.Text = TEST_TEXT;

            // Assert
            var lastPara = textFrame.Paragraphs.Last(); 
            lastPara.Text.Should().BeEquivalentTo(TEST_TEXT);
            textFrame.Paragraphs.Should().HaveCountGreaterThan(originParagraphsCount);

            pres.SaveAs(mStream);
            pres = SCPresentation.Open(mStream);
            textFrame = ((IAutoShape)pres.Slides[0].Shapes.First(sp => sp.Id == 4)).TextFrame;
            textFrame.Paragraphs.Last().Text.Should().BeEquivalentTo(TEST_TEXT);
            textFrame.Paragraphs.Should().HaveCountGreaterThan(originParagraphsCount);
        }

        [Fact]
        public void Paragraphs_Add_adds_paragraph()
        {
            // Arrange
            var pptxStream = GetTestStream("autoshape-case007.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var paragraphs = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 1").TextFrame.Paragraphs;
            
            // Act
            paragraphs.Add();
            
            // Assert
            paragraphs.Should().HaveCount(6);
        }

        [Fact]
        public void Paragraphs_Add_adds_new_text_paragraph_at_the_end_And_returns_added_paragraph_When_it_has_been_added_after_text_frame_changed()
        {
            var pres = SCPresentation.Open(Properties.Resources._001);
            var autoShape = (IAutoShape)pres.Slides[0].Shapes.First(sp => sp.Id == 3);
            var textBox = autoShape.TextFrame;
            var paragraphs = textBox.Paragraphs;
            var paragraph = textBox.Paragraphs.First();

            // Act
            textBox.Text = "A new text";
            var newParagraph = paragraphs.Add();

            // Assert
            newParagraph.Should().NotBeNull();
        }

        [Fact]
        public void CanTextChange_returns_false()
        {
            // Arrange
            var pptxStream = GetTestStream("autoshape-case006_field.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var textFrame = pres.Slides[0].Shapes.GetByName<IAutoShape>("Field 1").TextFrame;
            
            // Act
            var canTextChange = textFrame.CanChangeText();
            
            // Assert
            canTextChange.Should().BeFalse();
        }
    }
}

#endif