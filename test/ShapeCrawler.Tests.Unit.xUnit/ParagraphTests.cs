using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Helpers.Attributes;
using Xunit;

namespace ShapeCrawler.Tests.Unit;

[SuppressMessage("Usage", "xUnit1013:Public method should be marked as test")]
public class ParagraphTests : SCTest
{
    [Xunit.Theory]
    [SlideParagraphData("autoshape-case003.pptx", 1, "AutoShape 5", 1, 1)]
    [SlideParagraphData("autoshape-case003.pptx", 1, "AutoShape 5", 2, 2)]
    public void IndentLevel_Getter_returns_indent_level(IParagraph paragraph, int expectedLevel)
    {
        // Act
        var indentLevel = paragraph.IndentLevel;

        // Assert
        indentLevel.Should().Be(expectedLevel);
    }

    [Xunit.Theory]
    [MemberData(nameof(TestCasesAlignmentGetter))]
    public void Alignment_Getter_returns_text_alignment(IShape autoShape, TextAlignment expectedAlignment)
    {
        // Arrange
        var paragraph = autoShape.TextFrame.Paragraphs[0];

        // Act
        var textAlignment = paragraph.Alignment;

        // Assert
        textAlignment.Should().Be(expectedAlignment);
    }
    
    public static IEnumerable<object[]> TestCasesAlignmentGetter()
    {
        var pptx = StreamOf("001.pptx");
        var autoShape1 = new Presentation(pptx).Slides[0].Shapes.GetByName<IShape>("TextBox 3");
        yield return new object[] { autoShape1, TextAlignment.Center };

        var pptxStream2 = StreamOf("001.pptx");
        var pres2 = new Presentation(pptxStream2);
        var autoShape2 = pres2.Slides[0].Shapes.GetByName<IShape>("Head 1");
        yield return new object[] { autoShape2, TextAlignment.Center };
    }

    [Xunit.Theory]
    [MemberData(nameof(TestCasesParagraphsAlignmentSetter))]
    public void Alignment_Setter_updates_text_alignment(TestCase testCase)
    {
        // Arrange
        var pres = testCase.Presentation;
        var paragraph = testCase.AutoShape.TextFrame.Paragraphs[0];
        var mStream = new MemoryStream();

        // Act
        paragraph.Alignment = TextAlignment.Right;

        // Assert
        paragraph.Alignment.Should().Be(TextAlignment.Right);

        pres.SaveAs(mStream);
        testCase.SetPresentation(mStream);
        paragraph = testCase.AutoShape.TextFrame.Paragraphs[0];
        paragraph.Alignment.Should().Be(TextAlignment.Right);
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
    
    [Xunit.Theory]
    [MemberData(nameof(TestCasesParagraphText))]
    public void Text_Setter_sets_paragraph_text(TestElementQuery paragraphQuery, string newText, int expectedPortionsCount)
    {
        // Arrange
        var paragraph = paragraphQuery.GetParagraph();
        var mStream = new MemoryStream();
        var pres = paragraphQuery.Presentation;

        // Act
        paragraph.Text = newText;

        // Assert
        paragraph.Text.Should().BeEquivalentTo(newText);
        paragraph.Portions.Count.Should().Be(expectedPortionsCount);

        pres.SaveAs(mStream);
        paragraphQuery.Presentation = new Presentation(mStream);
        paragraph = paragraphQuery.GetParagraph();
        paragraph.Text.Should().BeEquivalentTo(newText);
        
        paragraph.Portions.Count.Should().Be(expectedPortionsCount);
    }
    
    public static IEnumerable<object[]> TestCasesParagraphText()
    {
        var paragraphQuery = new TestElementQuery
        {
            SlideIndex = 1,
            ShapeId = 4,
            ParagraphIndex = 2
        };
        paragraphQuery.Presentation = new Presentation(StreamOf("002.pptx"));
        yield return new object[] { paragraphQuery, "Text", 1 };
        
        var paragraphQuery2 = new TestElementQuery
        {
            SlideIndex = 1,
            ShapeId = 4,
            ParagraphIndex = 2
        };
        paragraphQuery2.Presentation = new Presentation(StreamOf("002.pptx"));
        yield return new object[] { paragraphQuery2, $"Text{Environment.NewLine}", 2 };
        
        var paragraphQuery3 = new TestElementQuery
        {
            SlideIndex = 1,
            ShapeId = 4,
            ParagraphIndex = 2
        };
        paragraphQuery3.Presentation = new Presentation(StreamOf("002.pptx"));
        yield return new object[] { paragraphQuery3, $"Text{Environment.NewLine}Text2", 3 };
        
        var paragraphQuery4 = new TestElementQuery
        {
            SlideIndex = 1,
            ShapeId = 4,
            ParagraphIndex = 2
        };
        paragraphQuery4.Presentation = new Presentation(StreamOf("002.pptx"));
        yield return new object[] { paragraphQuery4, $"Text{Environment.NewLine}Text2{Environment.NewLine}", 4 };
    }

    [Xunit.Theory]
    [SlideShapeData("autoshape-grouping.pptx", 1, "TextBox 5", 1.0)]
    [SlideShapeData("autoshape-grouping.pptx", 1, "TextBox 4", 1.5)]
    [SlideShapeData("autoshape-grouping.pptx", 1, "TextBox 3", 2.0)]
    public void Paragraph_Spacing_LineSpacingLines_returns_line_spacing_in_Lines(IShape shape, double expectedLines)
    {
        // Arrange
        var autoShape = (IShape)shape;
        var paragraph = autoShape.TextFrame!.Paragraphs[0];
            
        // Act
        var spacingLines = paragraph.Spacing.LineSpacingLines;
            
        // Assert
        spacingLines.Should().Be(expectedLines);
        paragraph.Spacing.LineSpacingPoints.Should().BeNull();
    }
        
    [Xunit.Theory]
    [SlideShapeData("autoshape-grouping.pptx", 1, "TextBox 6", 21.6)]
    public void Paragraph_Spacing_LineSpacingPoints_returns_line_spacing_in_Points(IShape shape, double expectedPoints)
    {
        // Arrange
        var autoShape = (IShape)shape;
        var paragraph = autoShape.TextFrame!.Paragraphs[0];
            
        // Act
        var spacingPoints = paragraph.Spacing.LineSpacingPoints;
            
        // Assert
        spacingPoints.Should().Be(expectedPoints);
        paragraph.Spacing.LineSpacingLines.Should().BeNull();
    }
}