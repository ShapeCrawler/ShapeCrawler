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