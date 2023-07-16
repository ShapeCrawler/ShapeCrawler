using System.Collections.Generic;
using System.IO;
using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Helpers.Attributes;
using Xunit;

// ReSharper disable SuggestVarOrType_SimpleTypes

namespace ShapeCrawler.Tests.Unit;

public class FontTests : SCTest
{
    [Test]
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
    
    [Test]
    public void Size_Getter_returns_font_size_of_non_first_portion()
    {
        // Arrange
        var pptx15 = GetInputStream("015.pptx");
        var pres15 = SCPresentation.Open(pptx15);
        var font1 = pres15.Slides[0].Shapes.GetById<IAutoShape>(5).TextFrame!.Paragraphs[0].Portions[2].Font;
        var font2 = SCPresentation.Open(GetInputStream("009_table.pptx")).Slides[2].Shapes.GetById<IAutoShape>(2).TextFrame!.Paragraphs[0].Portions[1].Font;

        // Act
        var fontSize1 = font1.Size;
        var fontSize2 = font2.Size;
        
        // Assert
        fontSize1.Should().Be(18);
        fontSize2.Should().Be(20);
    }

    [Test]
    public void Size_Getter_returns_Font_Size_of_Non_Placeholder_Table()
    {
        // Arrange
        var table = (ITable)SCPresentation.Open(GetInputStream("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 3);
        var cellPortion = table.Rows[0].Cells[0].TextFrame.Paragraphs[0].Portions[0];

        // Act-Assert
        cellPortion.Font.Size.Should().Be(18);
    }

    [Test]
    public void IsBold_GetterReturnsTrue_WhenFontOfNonPlaceholderTextIsBold()
    {
        // Arrange
        var nonPlaceholderAutoShapeCase1 =
            (IAutoShape)SCPresentation.Open(GetInputStream("020.pptx")).Slides[0].Shapes.First(sp => sp.Id == 3);
        IFont fontC1 = nonPlaceholderAutoShapeCase1.TextFrame.Paragraphs[0].Portions[0].Font;

        // Act-Assert
        fontC1.IsBold.Should().BeTrue();
    }

    [Test]
    public void IsBold_GetterReturnsTrue_WhenFontOfPlaceholderTextIsBold()
    {
        // Arrange
        IAutoShape placeholderAutoShape = (IAutoShape)SCPresentation.Open(GetInputStream("020.pptx")).Slides[1].Shapes.First(sp => sp.Id == 6);
        IPortion portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act
        bool isBold = portion.Font.IsBold;

        // Assert
        isBold.Should().BeTrue();
    }

    [Test]
    public void IsBold_GetterReturnsFalse_WhenFontOfNonPlaceholderTextIsNotBold()
    {
        // Arrange
        IAutoShape nonPlaceholderAutoShape = (IAutoShape)SCPresentation.Open(GetInputStream("020.pptx")).Slides[0].Shapes.First(sp => sp.Id == 2);
        IPortion portion = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act
        bool isBold = portion.Font.IsBold;

        // Assert
        isBold.Should().BeFalse();
    }

    [Test]
    public void IsBold_GetterReturnsFalse_WhenFontOfPlaceholderTextIsNotBold()
    {
        // Arrange
        var placeholderAutoShape = (IAutoShape)SCPresentation.Open(GetInputStream("020.pptx")).Slides[2].Shapes.First(sp => sp.Id == 7);
        var portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act
        bool isBold = portion.Font.IsBold;

        // Assert
        isBold.Should().BeFalse();
    }

    [Test]
    public void IsBold_Setter_AddsBoldForNonPlaceholderTextFont()
    {
        // Arrange
        var mStream = new MemoryStream();
        var pres20 = SCPresentation.Open(GetInputStream("020.pptx"));
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

    [Test]
    public void IsItalic_GetterReturnsTrue_WhenFontOfNonPlaceholderTextIsItalic()
    {
        // Arrange
        IAutoShape nonPlaceholderAutoShape = (IAutoShape)SCPresentation.Open(GetInputStream("020.pptx")).Slides[0].Shapes.First(sp => sp.Id == 3);
        IFont font = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0].Font;

        // Act
        bool isItalicFont = font.IsItalic;

        // Assert
        isItalicFont.Should().BeTrue();
    }

    [Test]
    public void IsItalic_GetterReturnsTrue_WhenFontOfPlaceholderTextIsItalic()
    {
        // Arrange
        IAutoShape placeholderAutoShape = (IAutoShape)SCPresentation.Open(GetInputStream("020.pptx")).Slides[2].Shapes.First(sp => sp.Id == 7);
        IPortion portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act-Assert
        portion.Font.IsItalic.Should().BeTrue();
    }

    [Test]
    public void IsItalic_Setter_SetsItalicFontForForNonPlaceholderText()
    {
        // Arrange
        var mStream = new MemoryStream();
        IPresentation presentation = SCPresentation.Open(GetInputStream("020.pptx"));
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

    [Test]
    public void IsItalic_SetterSetsNonItalicFontForPlaceholderText_WhenFalseValueIsPassed()
    {
        // Arrange
        var mStream = new MemoryStream();
        IPresentation presentation = SCPresentation.Open(GetInputStream("020.pptx"));
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

    [Test]
    public void Underline_SetUnderlineFont_WhenValueEqualsSetPassed()
    {
        // Arrange
        var mStream = new MemoryStream();
        IPresentation presentation = SCPresentation.Open(GetInputStream("020.pptx"));
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
}