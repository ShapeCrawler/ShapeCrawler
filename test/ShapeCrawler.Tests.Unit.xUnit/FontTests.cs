using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Helpers.Attributes;
using Xunit;

// ReSharper disable SuggestVarOrType_SimpleTypes

namespace ShapeCrawler.Tests.Unit;

public class FontTests : SCTest
{
    [Theory]
    [SlideShapeData("002.pptx", 2, 3, "Palatino Linotype")]
    [SlideShapeData("001.pptx", 1, 4, "Broadway")]
    [SlideShapeData("001.pptx", 1, 7, "Calibri Light")]
    [SlideShapeData("001.pptx", 5, 5, "Calibri Light")]
    [SlideShapeData("autoshape-grouping.pptx", 1, "Title 1", "Franklin Gothic Medium")]
    public void LatinName_Getter_returns_font_for_Latin_characters(IShape shape, string expectedFontName)
    {
        // Arrange
        var autoShape = shape;
        var font = autoShape.TextFrame!.Paragraphs[0].Portions[0].Font;

        // Act
        var fontName = font.LatinName;

        // Assert
        fontName.Should().Be(expectedFontName);
    }
    
    [Theory]
    [SlideShapeData("autoshape-grouping.pptx", 1, "TextBox 7", "SimSun")]
    public void EastAsianName_Getter_returns_font_for_East_Asian_characters(IShape shape, string expectedFontName)
    {
        // Arrange
        var autoShape =  shape;
        var font = autoShape.TextFrame!.Paragraphs[0].Portions[0].Font;

        // Act
        var fontName = font.EastAsianName;

        // Assert
        fontName.Should().Be(expectedFontName);
    }
    
    [Theory]
    [SlideShapeData("001.pptx", 1, "TextBox 3")]
    [SlideShapeData("001.pptx", 3, "Text Placeholder 3")]
    public void LatinName_Setter_sets_font_for_the_latin_characters(IShape shape)
    {
        // Arrange
        var autoShape =  shape;
        var font = autoShape.TextFrame!.Paragraphs[0].Portions[0].Font;

        // Act
        font.LatinName = "Time New Roman";

        // Assert
        font.LatinName.Should().Be("Time New Roman");
    }

    [Theory]
    [MasterShapeData("001.pptx", "Freeform: Shape 7", 18)]
    [SlideShapeData("020.pptx", 1, 3, 18)]
    [SlideShapeData("015.pptx", 2, 61, 18)]
    [SlideShapeData("009_table.pptx", 3, 2, 18)]
    [SlideShapeData("009_table.pptx", 4, 2, 44)]
    [SlideShapeData("009_table.pptx", 4, 3, 32)]
    [SlideShapeData("019.pptx", 1, 4103, 18)]
    [SlideShapeData("019.pptx", 1, 2, 12)]
    [SlideShapeData("014.pptx", 2, 5, 21)]
    [SlideShapeData("012_title-placeholder.pptx", 1, "Title 1", 20)]
    [SlideShapeData("010.pptx", 1, 2, 15)]
    [SlideShapeData("014.pptx", 4, 5, 12)]
    [SlideShapeData("014.pptx", 5, 4, 12)]
    [SlideShapeData("014.pptx", 6, 52, 27)]
    [SlideShapeData("autoshape-case016.pptx", 1, "Text Placeholder 1", 28)]
    public void Size_Getter_returns_font_size(IShape shape, int expectedSize)
    {
        // Arrange
        var autoShape =  shape;
        var font = autoShape.TextFrame!.Paragraphs[0].Portions[0].Font;
        
        // Act
        var fontSize = font.Size;
        
        // Assert
        fontSize.Should().Be(expectedSize);
    }

    [Theory]
    [MemberData(nameof(TestCasesSizeGetter))]
    public void Size_Getter_returns_font_size_of_Placeholder(TestCase testCase)
    {
        // Arrange
        var font = testCase.AutoShape.TextFrame!.Paragraphs[0].Portions[0].Font;
        var expectedFontSize = testCase.ExpectedInt;

        // Act
        var fontSize = font.Size;

        // Assert
        fontSize.Should().Be(expectedFontSize);
    }

    public static IEnumerable<object[]> TestCasesSizeGetter
    {
        get
        {
            var testCase1 = new TestCase("#1");
            testCase1.PresentationName = "028.pptx";
            testCase1.SlideNumber = 1;
            testCase1.ShapeId = 4098;
            testCase1.ExpectedInt = 32;
            yield return new object[] { testCase1 };

            var testCase2 = new TestCase("#2");
            testCase2.PresentationName = "029.pptx";
            testCase2.SlideNumber = 1;
            testCase2.ShapeName = "Content Placeholder 2";
            testCase2.ExpectedInt = 25;
            yield return new object[] { testCase2 };
        }
    }

    [Theory]
    [MemberData(nameof(TestCasesSizeSetter))]
    public void Size_Setter_sets_font_size(TestCase testCase)
    {
        // Arrange
        var pres = testCase.Presentation;
        var font = testCase.AutoShape.TextFrame!.Paragraphs[0].Portions[0].Font;
        var mStream = new MemoryStream();
        var oldSize = font.Size;
        var newSize = oldSize + 2;

        // Act
        font.Size = newSize;

        // Assert
        ((Presentation)pres).Validate();
        pres.SaveAs(mStream);
        testCase.SetPresentation(mStream);
        font = testCase.AutoShape.TextFrame!.Paragraphs[0].Portions[0].Font;
        font.Size.Should().Be(newSize);
    }

    public static IEnumerable<object[]> TestCasesSizeSetter
    {
        get
        {
            var testCase1 = new TestCase("#1");
            testCase1.PresentationName = "001.pptx";
            testCase1.SlideNumber = 1;
            testCase1.ShapeName = "TextBox 3";
            yield return new object[] { testCase1 };
            
            var testCase2 = new TestCase("#2");
            testCase2.PresentationName = "026.pptx";
            testCase2.SlideNumber = 1;
            testCase2.ShapeName = "AutoShape 1";
            yield return new object[] { testCase2 };
            
            var testCase3 = new TestCase("#3");
            testCase3.PresentationName = "autoshape-case016.pptx";
            testCase3.SlideNumber = 1;
            testCase3.ShapeName = "Text Placeholder 1";
            yield return new object[] { testCase3 };
        }
    }
    
    [Theory]
    [MemberData(nameof(TestCasesIsBold))]
    public void IsBold_Setter_sets_the_placeholder_font_to_be_bold(TestElementQuery portionQuery)
    {
        // Arrange
        var stream = new MemoryStream();
        var pres = portionQuery.Presentation;
        var font = portionQuery.GetParagraphPortion().Font;

        // Act
        font.IsBold = true;

        // Assert
        font.IsBold.Should().BeTrue();

        pres.SaveAs(stream);
        pres = new Presentation(stream);
        portionQuery.Presentation = pres;
        font = portionQuery.GetParagraphPortion().Font;
        font.IsBold.Should().BeTrue();
    }

    public static IEnumerable<object[]> TestCasesIsBold()
    {
        TestElementQuery portionRequestCase1 = new();
        portionRequestCase1.Presentation = new Presentation(StreamOf("020.pptx"));
        portionRequestCase1.SlideIndex = 2;
        portionRequestCase1.ShapeId = 7;
        portionRequestCase1.ParagraphIndex = 0;
        portionRequestCase1.PortionIndex = 0;

        TestElementQuery portionRequestCase2 = new();
        portionRequestCase2.Presentation = new Presentation(StreamOf("026.pptx"));
        portionRequestCase2.SlideIndex = 0;
        portionRequestCase2.ShapeId = 128;
        portionRequestCase2.ParagraphIndex = 0;
        portionRequestCase2.PortionIndex = 0;

        var testCases = new List<object[]>
        {
            new object[] { portionRequestCase1 },
            new object[] { portionRequestCase2 }
        };

        return testCases;
    }

    [Theory]
    [MemberData(nameof(TestCasesOffsetGetter))]
    public void OffsetEffect_Getter_returns_offset_of_Text(TestCase testCase)
    {
        // Arrange
        var font = testCase.AutoShape.TextFrame!.Paragraphs[0].Portions[1].Font;
        var expectedOffsetSize = testCase.ExpectedInt;

        // Act
        var offsetSize = font.OffsetEffect;

        // Assert
        offsetSize.Should().Be(expectedOffsetSize);
    }

    public static IEnumerable<object[]> TestCasesOffsetGetter
    {
        get
        {
            var testCase1 = new TestCase("#1");
            testCase1.PresentationName = "autoshape-case010.pptx";
            testCase1.SlideNumber = 1;
            testCase1.ShapeId = 2;
            testCase1.ExpectedInt = 50;
            yield return new object[] { testCase1 };

            var testCase2 = new TestCase("#2");
            testCase2.PresentationName = "autoshape-case010.pptx";
            testCase2.SlideNumber = 2;
            testCase2.ShapeName = "Title 1";
            testCase2.ExpectedInt = -32;
            yield return new object[] { testCase2 };
        }
    }

    [Theory]
    [MemberData(nameof(TestCasesOffsetSetter))]
    public void OffsetEffect_Setter_changes_Offset_of_paragraph_portion(TestCase testCase)
    {
        // Arrange
        var pres = testCase.Presentation;
        var font = testCase.AutoShape.TextFrame!.Paragraphs[0].Portions[0].Font;
        int superScriptOffset = testCase.ExpectedInt;
        var mStream = new MemoryStream();
        var oldOffsetSize = font.OffsetEffect;

        // Act
        font.OffsetEffect = superScriptOffset;
        pres.SaveAs(mStream);

        // Assert
        testCase.SetPresentation(mStream);
        font = testCase.AutoShape.TextFrame!.Paragraphs[0].Portions[0].Font;
        font.OffsetEffect.Should().NotBe(oldOffsetSize);
        font.OffsetEffect.Should().Be(superScriptOffset);
    }

    public static IEnumerable<object[]> TestCasesOffsetSetter
    {
        get
        {
            var testCase1 = new TestCase("#1");
            testCase1.PresentationName = "autoshape-case010.pptx";
            testCase1.SlideNumber = 3;
            testCase1.ShapeId = 2;
            testCase1.ExpectedInt = 12;
            yield return new object[] { testCase1 };

            var testCase2 = new TestCase("#2");
            testCase2.PresentationName = "autoshape-case010.pptx";
            testCase2.SlideNumber = 4;
            testCase2.ShapeName = "Title 1";
            testCase2.ExpectedInt = -27;
            yield return new object[] { testCase2 };
        }
    }
}