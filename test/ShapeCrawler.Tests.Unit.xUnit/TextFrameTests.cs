using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Helpers.Attributes;
using Xunit;

// ReSharper disable All
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit.xUnit
{
    [SuppressMessage("Usage", "xUnit1013:Public method should be marked as test")]
    public class TextFrameTests : SCTest
    {
        
        [Xunit.Theory]
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

            testTextBoxQuery.Presentation = new Presentation(mStream);
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
                    Presentation = new Presentation(StreamOf("001.pptx")),
                    SlideIndex = 0,
                    ShapeId = 3
                };
                testCases.Add(case1);
                
                var case2 = new TestElementQuery
                {
                    Presentation = new Presentation(StreamOf("020.pptx")),
                    SlideIndex = 2,
                    ShapeId = 8
                };
                testCases.Add(case2);
                
                var case3 = new TestElementQuery
                {
                    Presentation = new Presentation(StreamOf("001.pptx")),
                    SlideNumber = 2,
                    ShapeName = "Header 1",
                };
                testCases.Add(case3);
                
                var case4 = new TestElementQuery
                {
                    Presentation = new Presentation(StreamOf("autoshape-case004_subtitle.pptx")),
                    SlideNumber = 1,
                    ShapeName = "Subtitle 1",
                };
                testCases.Add(case4);
                
                var case5 = new TestElementQuery
                {
                    Presentation = new Presentation(StreamOf("autoshape-case008_text-frame.pptx")),
                    SlideNumber = 1,
                    ShapeName = "AutoShape 1",
                };
                testCases.Add(case5);

                return testCases;
            }
        }

        [Xunit.Theory]
        [SlideShapeData("autoshape-case012.pptx", 1, "Shape 1")]
        public void Text_Setter(IShape shape)
        {
            // Arrange
            var autoShape = (IShape)shape;
            var textFrame = autoShape.TextFrame;

            // Act
            var text = textFrame.Text;
            textFrame.Text = "some text";
            
            // Assert
            textFrame.Text.Should().BeEquivalentTo("some text");
        }
        
        [Xunit.Theory]
        [SlideShapeData("autoshape-case013.pptx", 1, "AutoShape 1")]
        public void Text_Setter_sets_long_text(IShape shape)
        {
            // Arrange
            var autoShape = (IShape)shape;
            var textFrame = autoShape.TextFrame;

            // Act
            var text = textFrame.Text;
            textFrame.Text = "Some sentence. Some sentence";
            
            // Assert
            shape.Height.Should().Be(88);
        }

        [Xunit.Theory]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 6", false)]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 2", true)]
        [SlideShapeData("autoshape-case013.pptx", 1, "AutoShape 1", true)]
        public void TextWrapped_Getter_returns_value_indicating_whether_text_is_wrapped_in_shape(IShape shape, bool isTextWrapped)
        {
            // Arrange
            var autoShape = (IShape)shape;
            var textFrame = autoShape.TextFrame!;

            // Act
            var textWrapped = textFrame.TextWrapped;

            // Assert
            textWrapped.Should().Be(isTextWrapped);
        }
        
        [Xunit.Theory]
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
                var pptxStream1 = StreamOf("009_table.pptx");
                var pres1 = new Presentation(pptxStream1);
                var autoShape1 = pres1.Slides[2].Shapes.GetById<IShape>(2);
                var textBox1 = autoShape1.TextFrame;
                var testCase1 = new TestCase<ITextFrame, int>(1, textBox1, 1);
                yield return new object[] { testCase1 };

                var pptxStream2 = StreamOf("020.pptx");
                var pres2 = new Presentation(pptxStream2);
                var autoShape2 = pres2.Slides[2].Shapes.GetById<IShape>(8);
                var textBox2 = autoShape2.TextFrame;
                var testCase2 = new TestCase<ITextFrame, int>(2, textBox2, 2);
                yield return new object[] { testCase2 };

                var pptxStream3 = StreamOf("009_table.pptx");
                var pres3 = new Presentation(pptxStream3);
                var table3 = pres3.Slides[2].Shapes.GetById<ITable>(3);
                var textBox3 = table3.Rows[0].Cells[0].TextFrame;
                var testCase3 = new TestCase<ITextFrame, int>(3, textBox3, 2);
                yield return new object[] { testCase3 };

                var pptxStream4 = StreamOf("001.pptx");
                var pres4 = new Presentation(pptxStream4);
                var autoShape4 = pres4.Slides[1].Shapes.GetById<IShape>(2);
                var textBox4 = autoShape4.TextFrame;
                var testCase4 = new TestCase<ITextFrame, int>(4, textBox4, 1);
                yield return new object[] { testCase4 };
            }
        }

        [Xunit.Theory]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 2", 0.25)]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 3", 0.30)]
        public void LeftMargin_getter_returns_left_margin_of_text_frame_in_centimeters(IShape shape, double expectedMargin)
        {
            // Arrange
            var autoShape = (IShape)shape;
            var textFrame = autoShape.TextFrame;
            
            // Act
            var leftMargin = textFrame.LeftMargin;
            
            // Assert
            leftMargin.Should().Be(expectedMargin);
        }
        
        [Xunit.Theory]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 2")]
        public void LeftMargin_setter_sets_left_margin_of_text_frame_in_centimeters(IShape shape)
        {
            // Arrange
            var autoShape = (IShape)shape;
            var textFrame = autoShape.TextFrame;
            
            // Act
            textFrame.LeftMargin = 0.5;
            
            // Assert
            textFrame.LeftMargin.Should().Be(0.5);
        }
        
        [Xunit.Theory]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 2", 0.25)]
        public void RightMargin_getter_returns_right_margin_of_text_frame_in_centimeters(IShape shape, double expectedMargin)
        {
            // Arrange
            var autoShape = (IShape)shape;
            var textFrame = autoShape.TextFrame;
            
            // Act
            var rightMargin = textFrame.RightMargin;
            
            // Assert
            rightMargin.Should().Be(expectedMargin);
        }
        
        [Xunit.Theory]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 2", 0.13)]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 3", 0.14)]
        public void TopMargin_getter_returns_top_margin_of_text_frame_in_centimeters(IShape shape, double expectedMargin)
        {
            // Arrange
            var autoShape = (IShape)shape;
            var textFrame = autoShape.TextFrame;
            
            // Act
            var topMargin = textFrame.TopMargin;
            
            // Assert
            topMargin.Should().Be(expectedMargin);
        }
        
        [Xunit.Theory]
        [SlideShapeData("autoshape-case003.pptx", 1, "AutoShape 2", 0.13)]
        public void BottomMargin_getter_returns_bottom_margin_of_text_frame_in_centimeters(IShape shape, double expectedMargin)
        {
            // Arrange
            var autoShape = (IShape)shape;
            var textFrame = autoShape.TextFrame;
            
            // Act
            var bottomMargin = textFrame.BottomMargin;
            
            // Assert
            bottomMargin.Should().Be(expectedMargin);
        }

        [Xunit.Theory]
        [MemberData(nameof(TestCasesTextFrameXPath))]
        public void GetPresentationSlideTextFrameXPath(string presentationName, int slideNumber, string[] expectedXPath)
        {
            // Arrange
            var pres = new Presentation(StreamOf(presentationName));
            var slide = pres.Slides[slideNumber];
            var textFrame = slide.TextFrames();

            // Act
            var actualXPath = textFrame.Select(tf => tf.SDKXPath).ToArray();

            // Assert
            actualXPath.Should().BeEquivalentTo(expectedXPath);
        }

        public static IEnumerable<object[]> TestCasesTextFrameXPath()
        {
            yield return new object[]
            {
                "054_get_shape_xpath.pptx", 0,
                new string[]
                {
                    "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[1]/p:txBody[1]",
                    "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[2]/p:txBody[1]"
                }
            };
            yield return new object[]
            {
                "054_get_shape_xpath.pptx", 1,
                new string[]
                {
                    "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[1]/p:txBody[1]",
                    "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[2]/p:txBody[1]",
                    "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[3]/p:txBody[1]"
                }
            };
        }
    }
}