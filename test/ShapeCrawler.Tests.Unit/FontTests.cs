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
        var autoShape = (IAutoShape)shape;
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
        var autoShape = (IAutoShape)shape;
        var font = autoShape.TextFrame!.Paragraphs[0].Portions[0].Font;

        // Act
        var fontName = font.EastAsianName;

        // Assert
        fontName.Should().Be(expectedFontName);
    }
    
    [Theory]
    [SlideShapeData("001.pptx", 1, "TextBox 3")]
    public void EastAsianName_Setter_sets_font_for_the_east_asian_characters(IShape shape)
    {
        // Arrange
        var autoShape = (IAutoShape)shape;
        var font = autoShape.TextFrame!.Paragraphs[0].Portions[0].Font;

        // Act
        font.EastAsianName = "SimSun";

        // Assert
        font.EastAsianName.Should().Be("SimSun");
        var errors = PptxValidator.Validate(shape.SlideStructure.Presentation);
        errors.Should().BeEmpty();
    }
    
    [Fact]
    public void EastAsianName_Setter_sets_font_for_the_east_asian_characters_in_New_Presentation()
    {
        // Arrange
        var pres = SCPresentation.Create();
        var slide = pres.Slides[0];
        var rectangle = slide.Shapes.AddRectangle(10, 10, 10, 10);
        var font = rectangle.TextFrame!.Paragraphs[0].Portions[0].Font;

        // Act
        font.EastAsianName = "SimSun";

        // Assert
        font.EastAsianName.Should().Be("SimSun");
        var errors = PptxValidator.Validate(slide.Presentation);
        errors.Should().BeEmpty();
    }

    [Theory]
    [SlideShapeData("001.pptx", 1, "TextBox 3")]
    [SlideShapeData("001.pptx", 3, "Text Placeholder 3")]
    public void LatinName_Setter_sets_font_for_the_latin_characters(IShape shape)
    {
        // Arrange
        var autoShape = (IAutoShape)shape;
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
        var autoShape = (IAutoShape)shape;
        var font = autoShape.TextFrame!.Paragraphs[0].Portions[0].Font;
        
        // Act
        var fontSize = font.Size;
        
        // Assert
        fontSize.Should().Be(expectedSize);
    }
    
    [Fact]
    public void Size_Getter_returns_font_size_of_non_first_portion()
    {
        // Arrange
        var pptx15 = GetTestStream("015.pptx");
        var pres15 = SCPresentation.Open(pptx15);
        var font1 = pres15.Slides[0].Shapes.GetById<IAutoShape>(5).TextFrame!.Paragraphs[0].Portions[2].Font;
        var font2 = SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[2].Shapes.GetById<IAutoShape>(2).TextFrame!.Paragraphs[0].Portions[1].Font;

        // Act
        var fontSize1 = font1.Size;
        var fontSize2 = font2.Size;
        
        // Assert
        fontSize1.Should().Be(18);
        fontSize2.Should().Be(20);
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

    [Fact]
    public void Size_Getter_returns_Font_Size_of_Non_Placeholder_Table()
    {
        // Arrange
        var table = (ITable)SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 3);
        var cellPortion = table.Rows[0].Cells[0].TextFrame.Paragraphs[0].Portions[0];

        // Act-Assert
        cellPortion.Font.Size.Should().Be(18);
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
        var errors = PptxValidator.Validate(pres);
        errors.Should().BeEmpty();
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
    [SlideShapeData("#1", "001.pptx", 1, "TextBox 3")]
    [SlideShapeData("#2", "026.pptx", 1, "AutoShape 1")]
    [SlideShapeData("#3", "autoshape-case016.pptx", 1, "Text Placeholder 1")]
    public void CanChange_returns_true(string displayName, IShape shape)
    {
        // Arrange
        var autoShape = (IAutoShape)shape;
        var font = autoShape.TextFrame!.Paragraphs[0].Portions[0].Font;

        // Act
        var canChange = font.CanChange();

        // Assert
        canChange.Should().BeTrue();
    }
    
    [Fact]
    public void IsBold_GetterReturnsTrue_WhenFontOfNonPlaceholderTextIsBold()
    {
        // Arrange
        var nonPlaceholderAutoShapeCase1 =
            (IAutoShape)SCPresentation.Open(GetTestStream("020.pptx")).Slides[0].Shapes.First(sp => sp.Id == 3);
        IFont fontC1 = nonPlaceholderAutoShapeCase1.TextFrame.Paragraphs[0].Portions[0].Font;

        // Act-Assert
        fontC1.IsBold.Should().BeTrue();
    }

    [Fact]
    public void IsBold_GetterReturnsTrue_WhenFontOfPlaceholderTextIsBold()
    {
        // Arrange
        IAutoShape placeholderAutoShape = (IAutoShape)SCPresentation.Open(GetTestStream("020.pptx")).Slides[1].Shapes.First(sp => sp.Id == 6);
        IPortion portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act
        bool isBold = portion.Font.IsBold;

        // Assert
        isBold.Should().BeTrue();
    }

    [Fact]
    public void IsBold_GetterReturnsFalse_WhenFontOfNonPlaceholderTextIsNotBold()
    {
        // Arrange
        IAutoShape nonPlaceholderAutoShape = (IAutoShape)SCPresentation.Open(GetTestStream("020.pptx")).Slides[0].Shapes.First(sp => sp.Id == 2);
        IPortion portion = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act
        bool isBold = portion.Font.IsBold;

        // Assert
        isBold.Should().BeFalse();
    }

    [Fact]
    public void IsBold_GetterReturnsFalse_WhenFontOfPlaceholderTextIsNotBold()
    {
        // Arrange
        var placeholderAutoShape = (IAutoShape)SCPresentation.Open(GetTestStream("020.pptx")).Slides[2].Shapes.First(sp => sp.Id == 7);
        var portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act
        bool isBold = portion.Font.IsBold;

        // Assert
        isBold.Should().BeFalse();
    }

    [Fact]
    public void IsBold_Setter_AddsBoldForNonPlaceholderTextFont()
    {
        // Arrange
        var mStream = new MemoryStream();
        var pres20 = SCPresentation.Open(GetTestStream("020.pptx"));
        IPresentation presentation = pres20;
        IAutoShape nonPlaceholderAutoShape = (IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 2);
        IPortion portion = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act
        portion.Font.IsBold = true;

        // Assert
        portion.Font.IsBold.Should().BeTrue();
        presentation.SaveAs(mStream);
        presentation = SCPresentation.Open(mStream);
        nonPlaceholderAutoShape = (IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 2);
        portion = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0];
        portion.Font.IsBold.Should().BeTrue();
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
        pres = SCPresentation.Open(stream);
        portionQuery.Presentation = pres;
        font = portionQuery.GetParagraphPortion().Font;
        font.IsBold.Should().BeTrue();
    }

    public static IEnumerable<object[]> TestCasesIsBold()
    {
        TestElementQuery portionRequestCase1 = new();
        portionRequestCase1.Presentation = SCPresentation.Open(GetTestStream("020.pptx"));
        portionRequestCase1.SlideIndex = 2;
        portionRequestCase1.ShapeId = 7;
        portionRequestCase1.ParagraphIndex = 0;
        portionRequestCase1.PortionIndex = 0;

        TestElementQuery portionRequestCase2 = new();
        portionRequestCase2.Presentation = SCPresentation.Open(GetTestStream("026.pptx"));
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

    [Fact]
    public void IsItalic_GetterReturnsTrue_WhenFontOfNonPlaceholderTextIsItalic()
    {
        // Arrange
        IAutoShape nonPlaceholderAutoShape = (IAutoShape)SCPresentation.Open(GetTestStream("020.pptx")).Slides[0].Shapes.First(sp => sp.Id == 3);
        IFont font = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0].Font;

        // Act
        bool isItalicFont = font.IsItalic;

        // Assert
        isItalicFont.Should().BeTrue();
    }

    [Fact]
    public void IsItalic_GetterReturnsTrue_WhenFontOfPlaceholderTextIsItalic()
    {
        // Arrange
        IAutoShape placeholderAutoShape = (IAutoShape)SCPresentation.Open(GetTestStream("020.pptx")).Slides[2].Shapes.First(sp => sp.Id == 7);
        IPortion portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act-Assert
        portion.Font.IsItalic.Should().BeTrue();
    }

    [Fact]
    public void IsItalic_Setter_SetsItalicFontForForNonPlaceholderText()
    {
        // Arrange
        var mStream = new MemoryStream();
        IPresentation presentation = SCPresentation.Open(GetTestStream("020.pptx"));
        IAutoShape nonPlaceholderAutoShape = (IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 2);
        IPortion portion = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act
        portion.Font.IsItalic = true;

        // Assert
        portion.Font.IsItalic.Should().BeTrue();
        presentation.SaveAs(mStream);
        presentation = SCPresentation.Open(mStream);
        nonPlaceholderAutoShape = (IAutoShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 2);
        portion = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0];
        portion.Font.IsItalic.Should().BeTrue();
    }

    [Fact]
    public void IsItalic_SetterSetsNonItalicFontForPlaceholderText_WhenFalseValueIsPassed()
    {
        // Arrange
        var mStream = new MemoryStream();
        IPresentation presentation = SCPresentation.Open(GetTestStream("020.pptx"));
        IAutoShape placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        IPortion portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act
        portion.Font.IsItalic = false;

        // Assert
        portion.Font.IsItalic.Should().BeFalse();
        presentation.SaveAs(mStream);

        presentation = SCPresentation.Open(mStream);
        placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];
        portion.Font.IsItalic.Should().BeFalse();
    }

    [Fact]
    public void Underline_SetUnderlineFont_WhenValueEqualsSetPassed()
    {
        // Arrange
        var mStream = new MemoryStream();
        IPresentation presentation = SCPresentation.Open(GetTestStream("020.pptx"));
        IAutoShape placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        IPortion portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act
        portion.Font.Underline = DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Single;

        // Assert
        portion.Font.Underline.Should().Be(DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Single);
        presentation.SaveAs(mStream);

        presentation = SCPresentation.Open(mStream);
        placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];
        portion.Font.Underline.Should().Be(DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Single);
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