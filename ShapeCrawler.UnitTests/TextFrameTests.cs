#if DEBUG

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.UnitTests.Helpers;
using ShapeCrawler.UnitTests.Helpers.Attributes;
using ShapeCrawler.UnitTests.Helpers;
using Xunit;

// ReSharper disable All
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.UnitTests
{
    public class TextFrameTests : ShapeCrawlerTest
    {
        [Fact]
        public void Text_Getter_returns_text_of_table_Cell()
        {
            // Arrange
            var pptx8 = GetTestStream("008.pptx");
            var pres8 = SCPresentation.Open(pptx8);
            var pptx1 = Assets.GetStream("001.pptx");
            var pres1 = SCPresentation.Open(pptx1);
            var pptx9 = GetTestStream("009_table.pptx");
            var pres9 = SCPresentation.Open(pptx9);
            var textFrame1 = ((IAutoShape)SCPresentation.Open(GetTestStream("008.pptx")).Slides[0].Shapes.First(sp => sp.Id == 3)).TextFrame;
            var textFrame2 = ((ITable)SCPresentation.Open(Assets.GetStream("001.pptx")).Slides[1].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0]
                .TextFrame;
            var textFrame3 = ((ITable)SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0]
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
            var pptx = GetTestStream("autoshape-case005_text-frame.pptx");
            var pres = SCPresentation.Open(pptx);
            var textFrame = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 1").TextFrame;
            var modifiedPres = new MemoryStream();

            // Act
            textFrame.Text = textFrame.Text.Replace("{{replace_this}}", "confirm this");
            textFrame.Text = textFrame.Text.Replace("{{replace_that}}", "confirm that");

            // Assert
            pres.SaveAs(modifiedPres);
            pres.Close();
            pres = SCPresentation.Open(modifiedPres);
            textFrame = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 1").TextFrame;
            textFrame.Text.Should().ContainAll("confirm this", "confirm that");
        }
        
        [Theory]
        [MemberData(nameof(TestCasesTextSetter))]
        public void Text_Setter_updates_content(TestElementQuery testTextBoxQuery)
        {
            // Arrange
            var pres = testTextBoxQuery.Presentation;
            var textFrame = testTextBoxQuery.GetAutoShape().TextFrame;
            const string newText = "Test";
            var mStream = new MemoryStream();

            // Act
            textFrame.Text = newText;

            // Assert
            textFrame.Text.Should().BeEquivalentTo(newText);
            textFrame.Paragraphs.Should().HaveCount(1);

            pres.SaveAs(mStream);
            pres.Close();

            testTextBoxQuery.Presentation = SCPresentation.Open(mStream);
            textFrame = testTextBoxQuery.GetAutoShape().TextFrame;
            textFrame.Text.Should().BeEquivalentTo(newText);
            textFrame.Paragraphs.Should().HaveCount(1);
        }
        
        public static TheoryData<TestElementQuery> TestCasesTextSetter
        {
            get
            {
                var testCases = new TheoryData<TestElementQuery>();
                
                var case1 = new TestElementQuery
                {
                    Presentation = SCPresentation.Open(Assets.GetStream("001.pptx")),
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
                    Presentation = SCPresentation.Open(Assets.GetStream("001.pptx")),
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
            var pptxStream = Assets.GetStream("001.pptx");
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
        public void Text_Setter_resizes_shape_to_fit_text()
        {
            // Arrange
            var pptxStream = GetTestStream("autoshape-case003.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var shape = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 4");
            var textFrame = shape.TextFrame;

            // Act
            textFrame.Text = "AutoShape 4 some text";

            // Assert
            shape.Height.Should().Be(46);
            shape.Y.Should().Be(152);
            var errors = PptxValidator.Validate(shape.SlideObject.Presentation);
            errors.Should().BeEmpty();
        }
        
        [Theory]
        [SlideShapeData("autoshape-case012.pptx", 1, "Shape 1")]
        public void Text_Setter_should_not_throw_exception(IShape shape)
        {
            // Arrange
            var autoShape = (IAutoShape)shape;
            var textFrame = autoShape.TextFrame;

            // Act-Assert
            var text = textFrame.Text;
            textFrame.Text = "some text";
        }
        
        [Theory]
        [SlideShapeData("autoshape-case013.pptx", 1, "AutoShape 1")]
        public void Text_Setter_sets_long_text(IShape shape)
        {
            // Arrange
            var autoShape = (IAutoShape)shape;
            var textFrame = autoShape.TextFrame;

            // Act
            var text = textFrame.Text;
            textFrame.Text = "Some sentence. Some sentence";
            
            // Assert
            shape.Height.Should().Be(88);
        }
        
        [Fact]
        public void Text_Setter_sets_text_for_New_Shape()
        {
            // Arrange
            var pres = SCPresentation.Create();
            var shapes = pres.Slides[0].Shapes;
            var autoShape = shapes.AutoShapes.AddRectangle( 50, 60, 100, 70);
            var textFrame = autoShape.TextFrame!;
            
            // Act
            textFrame.Text = "Test";
    
            // Assert
            textFrame.Text.Should().Be("Test");
            var errors = PptxValidator.Validate(pres);
            errors.Should().BeEmpty();
        }

        [Theory]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 6", false)]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 2", true)]
        [SlideShapeData("autoshape-case013.pptx", 1, "AutoShape 1", true)]
        public void TextWrapped_Getter_returns_value_indicating_whether_text_is_wrapped_in_shape(IShape shape, bool isTextWrapped)
        {
            // Arrange
            var autoShape = (IAutoShape)shape;
            var textFrame = autoShape.TextFrame!;

            // Act
            var textWrapped = textFrame.TextWrapped;

            // Assert
            textWrapped.Should().Be(isTextWrapped);
        }
        
        [Fact]
        public void AutofitType_Setter_resizes_width()
        {
            // Arrange
            var pptxStream = GetTestStream("autoshape-case003.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var shape = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 6");
            var textFrame = shape.TextFrame!;

            // Act
            textFrame.AutofitType = SCAutofitType.Resize;

            // Assert
            shape.Width.Should().Be(107);
            var errors = PptxValidator.Validate(pres);
            errors.Should().BeEmpty();
        }
        
        [Theory]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 7")]
        [SlideShapeData("001.pptx", 1, "Head 1")]
        [SlideShapeData("autoshape-case014.pptx", 1, "Content Placeholder 1")]
        public void AutofitType_Setter_sets_autofit_type(IShape shape)
        {
            // Arrange
            var autoShape = (IAutoShape)shape;
            var textFrame = autoShape.TextFrame!;

            // Act
            textFrame.AutofitType = SCAutofitType.Resize;

            // Assert
            textFrame.AutofitType.Should().Be(SCAutofitType.Resize);
            var errors = PptxValidator.Validate(shape.SlideObject.Presentation);
            errors.Should().BeEmpty();
        }
        
        [Fact]
        public void AutofitType_Setter_updates_height()
        {
            // Arrange
            var pptxStream = GetTestStream("autoshape-case003.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var shape = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 7");
            var textFrame = shape.TextFrame!;

            // Act
            textFrame.AutofitType = SCAutofitType.Resize;

            // Assert
            shape.Height.Should().Be(35);
            var errors = PptxValidator.Validate(pres);
            errors.Should().BeEmpty();
        }
        
        [Fact]
        public void AutofitType_Getter_returns_text_autofit_type()
        {
            // Arrange
            var pptx = Assets.GetStream("001.pptx");
            var pres = SCPresentation.Open(pptx);
            var autoShape = pres.Slides[0].Shapes.GetById<IAutoShape>(9);
            var textBox = autoShape.TextFrame;

            // Act
            var autofitType = textBox.AutofitType;

            // Assert
            autofitType.Should().Be(SCAutofitType.Shrink);
        }

        [Fact]
        public void Shape_IsAutoShape()
        {
            // Arrange
            var pres8 = SCPresentation.Open(GetTestStream("008.pptx"));
            var pres21 = SCPresentation.Open(GetTestStream("021.pptx"));
            IShape shapeCase1 = SCPresentation.Open(GetTestStream("008.pptx")).Slides[0].Shapes.First(sp => sp.Id == 3);
            IShape shapeCase2 = SCPresentation.Open(GetTestStream("021.pptx")).Slides[3].Shapes.First(sp => sp.Id == 2);
            IShape shapeCase3 = SCPresentation.Open(GetTestStream("011_dt.pptx")).Slides[0].Shapes.First(sp => sp.Id == 54275);

            // Act
            var autoShapeCase1 = shapeCase1 as IAutoShape;
            var autoShapeCase2 = shapeCase2 as IAutoShape;
            var autoShapeCase3 = shapeCase3 as IAutoShape;

            // Assert
            autoShapeCase1.Should().NotBeNull();
            autoShapeCase2.Should().NotBeNull();
            autoShapeCase3.Should().NotBeNull();
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

                var pptxStream4 = Assets.GetStream("001.pptx");
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
            var pres = SCPresentation.Open(Assets.GetStream("001.pptx"));
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
            var pres = SCPresentation.Open(Assets.GetStream("001.pptx"));
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

        [Theory]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 2", 0.25)]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 3", 0.30)]
        public void LeftMargin_getter_returns_left_margin_of_text_frame_in_centimeters(IShape shape, double expectedMargin)
        {
            // Arrange
            var autoShape = (IAutoShape)shape;
            var textFrame = autoShape.TextFrame;
            
            // Act
            var leftMargin = textFrame.LeftMargin;
            
            // Assert
            leftMargin.Should().Be(expectedMargin);
        }
        
        [Theory]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 2")]
        public void LeftMargin_setter_sets_left_margin_of_text_frame_in_centimeters(IShape shape)
        {
            // Arrange
            var autoShape = (IAutoShape)shape;
            var textFrame = autoShape.TextFrame;
            
            // Act
            textFrame.LeftMargin = 0.5;
            
            // Assert
            textFrame.LeftMargin.Should().Be(0.5);
        }
        
        [Theory]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 2", 0.25)]
        public void RightMargin_getter_returns_right_margin_of_text_frame_in_centimeters(IShape shape, double expectedMargin)
        {
            // Arrange
            var autoShape = (IAutoShape)shape;
            var textFrame = autoShape.TextFrame;
            
            // Act
            var rightMargin = textFrame.RightMargin;
            
            // Assert
            rightMargin.Should().Be(expectedMargin);
        }
        
        [Theory]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 2", 0.13)]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 3", 0.14)]
        public void TopMargin_getter_returns_top_margin_of_text_frame_in_centimeters(IShape shape, double expectedMargin)
        {
            // Arrange
            var autoShape = (IAutoShape)shape;
            var textFrame = autoShape.TextFrame;
            
            // Act
            var topMargin = textFrame.TopMargin;
            
            // Assert
            topMargin.Should().Be(expectedMargin);
        }
        
        [Theory]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 2", 0.13)]
        public void BottomMargin_getter_returns_bottom_margin_of_text_frame_in_centimeters(IShape shape, double expectedMargin)
        {
            // Arrange
            var autoShape = (IAutoShape)shape;
            var textFrame = autoShape.TextFrame;
            
            // Act
            var bottomMargin = textFrame.BottomMargin;
            
            // Assert
            bottomMargin.Should().Be(expectedMargin);
        }
    }
}

#endif