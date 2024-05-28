using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Helpers.Attributes;

// ReSharper disable SuggestVarOrType_SimpleTypes

namespace ShapeCrawler.Tests.Unit;

public class FontTests : SCTest
{
    [Test]
    public void EastAsianName_Setter_sets_font_for_the_east_asian_characters_in_New_Presentation()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.Shapes.AddRectangle(10, 10, 10, 10);
        var rectangle = slide.Shapes.Last();
        rectangle.TextFrame.Paragraphs[0].Portions.AddText("test");
        var font = rectangle.TextFrame!.Paragraphs[0].Portions[0].Font;

        // Act
        font.EastAsianName = "SimSun";

        // Assert
        font.EastAsianName.Should().Be("SimSun");
        pres.Validate();
    }
    
    [Test]
    public void Size_Getter_returns_font_size_of_non_first_portion()
    {
        // Arrange
        var pres1 = new Presentation(StreamOf("015.pptx"));
        var font1 = pres1.Slides[0].Shapes.GetById<IShape>(5).TextFrame!.Paragraphs[0].Portions[2].Font;
        var pres2 = new Presentation(StreamOf("009_table.pptx"));
        var font2 = pres2.Slides[2].Shapes.GetById<IShape>(2).TextFrame!.Paragraphs[0].Portions[1].Font;

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
        var pres = new Presentation(StreamOf("009_table.pptx"));
        var table = pres.Slides[2].Shapes.GetById<ITable>(3);
        var portion = table.Rows[0].Cells[0].TextFrame.Paragraphs[0].Portions[0];

        // Act-Assert
        portion.Font.Size.Should().Be(18);
    }

    [Test]
    public void IsBold_GetterReturnsTrue_WhenFontOfNonPlaceholderTextIsBold()
    {
        // Arrange
        var nonPlaceholderAutoShapeCase1 =
            (IShape)new Presentation(StreamOf("020.pptx")).Slides[0].Shapes.First(sp => sp.Id == 3);
        ITextPortionFont fontC1 = nonPlaceholderAutoShapeCase1.TextFrame.Paragraphs[0].Portions[0].Font;

        // Act-Assert
        fontC1.IsBold.Should().BeTrue();
    }

    [Test]
    public void IsBold_GetterReturnsTrue_WhenFontOfPlaceholderTextIsBold()
    {
        // Arrange
        var pres = new Presentation(StreamOf("020.pptx"));
        var placeholderAutoShape = pres.Slides[1].Shapes.GetById<IShape>(6);
        var portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act
        var isBold = portion.Font.IsBold;

        // Assert
        isBold.Should().BeTrue();
    }

    [Test]
    public void IsBold_GetterReturnsFalse_WhenFontOfNonPlaceholderTextIsNotBold()
    {
        // Arrange
        IShape nonPlaceholderAutoShape = (IShape)new Presentation(StreamOf("020.pptx")).Slides[0].Shapes.First(sp => sp.Id == 2);
        IParagraphPortion portion = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act
        bool isBold = portion.Font.IsBold;

        // Assert
        isBold.Should().BeFalse();
    }

    [Test]
    public void IsBold_GetterReturnsFalse_WhenFontOfPlaceholderTextIsNotBold()
    {
        // Arrange
        var placeholderAutoShape = new Presentation(StreamOf("020.pptx")).Slides[2].Shapes.First(sp => sp.Id == 7);
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
        var pres20 = new Presentation(StreamOf("020.pptx"));
        IPresentation presentation = pres20;
        IShape nonPlaceholderAutoShape = (IShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 2);
        IParagraphPortion portion = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act
        portion.Font.IsBold = true;

        // Assert
        portion.Font.IsBold.Should().BeTrue();
        presentation.SaveAs(mStream);
        presentation = new Presentation(mStream);
        nonPlaceholderAutoShape = (IShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 2);
        portion = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0];
        portion.Font.IsBold.Should().BeTrue();
    }

    [Test]
    public void IsItalic_GetterReturnsTrue_WhenFontOfNonPlaceholderTextIsItalic()
    {
        // Arrange
        IShape nonPlaceholderAutoShape = (IShape)new Presentation(StreamOf("020.pptx")).Slides[0].Shapes.First(sp => sp.Id == 3);
        ITextPortionFont font = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0].Font;

        // Act
        bool isItalicFont = font.IsItalic;

        // Assert
        isItalicFont.Should().BeTrue();
    }

    [Test]
    public void IsItalic_GetterReturnsTrue_WhenFontOfPlaceholderTextIsItalic()
    {
        // Arrange
        IShape placeholderAutoShape = (IShape)new Presentation(StreamOf("020.pptx")).Slides[2].Shapes.First(sp => sp.Id == 7);
        IParagraphPortion portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act-Assert
        portion.Font.IsItalic.Should().BeTrue();
    }

    [Test]
    public void IsItalic_Setter_SetsItalicFontForForNonPlaceholderText()
    {
        // Arrange
        var mStream = new MemoryStream();
        IPresentation presentation = new Presentation(StreamOf("020.pptx"));
        IShape nonPlaceholderAutoShape = (IShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 2);
        IParagraphPortion portion = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act
        portion.Font.IsItalic = true;

        // Assert
        portion.Font.IsItalic.Should().BeTrue();
        presentation.SaveAs(mStream);
        presentation = new Presentation(mStream);
        nonPlaceholderAutoShape = (IShape)presentation.Slides[0].Shapes.First(sp => sp.Id == 2);
        portion = nonPlaceholderAutoShape.TextFrame.Paragraphs[0].Portions[0];
        portion.Font.IsItalic.Should().BeTrue();
    }

    [Test]
    public void IsItalic_SetterSetsNonItalicFontForPlaceholderText_WhenFalseValueIsPassed()
    {
        // Arrange
        var mStream = new MemoryStream();
        IPresentation presentation = new Presentation(StreamOf("020.pptx"));
        IShape placeholderAutoShape = (IShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        IParagraphPortion portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act
        portion.Font.IsItalic = false;

        // Assert
        portion.Font.IsItalic.Should().BeFalse();
        presentation.SaveAs(mStream);

        presentation = new Presentation(mStream);
        placeholderAutoShape = (IShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];
        portion.Font.IsItalic.Should().BeFalse();
    }

    [Test]
    public void Underline_SetUnderlineFont_WhenValueEqualsSetPassed()
    {
        // Arrange
        var mStream = new MemoryStream();
        IPresentation presentation = new Presentation(StreamOf("020.pptx"));
        IShape placeholderAutoShape = (IShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        IParagraphPortion portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act
        portion.Font.Underline = DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Single;

        // Assert
        portion.Font.Underline.Should().Be(DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Single);
        presentation.SaveAs(mStream);

        presentation = new Presentation(mStream);
        placeholderAutoShape = (IShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        portion = placeholderAutoShape.TextFrame.Paragraphs[0].Portions[0];
        portion.Font.Underline.Should().Be(DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Single);
    }
    
    [Test]
    [TestCase("001.pptx", 1, "TextBox 3")]
    public void EastAsianName_Setter_sets_font_for_the_east_asian_characters(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(StreamOf(file));
        var shape = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName);
        var font = shape.TextFrame.Paragraphs[0].Portions[0].Font;

        // Act
        font.EastAsianName = "SimSun";

        // Assert
        font.EastAsianName.Should().Be("SimSun");
        pres.Validate();
    }
    
    [Test]
    [SlideShape("002.pptx", 2, 3, "Palatino Linotype")]
    [SlideShape("001.pptx", 1, 4, "Broadway")]
    [SlideShape("001.pptx", 1, 7, "Calibri Light")]
    [SlideShape("001.pptx", 5, 5, "Calibri Light")]
    [SlideShape("autoshape-grouping.pptx", 1, "Title 1", "Franklin Gothic Medium")]
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
    
    [Test]
    [SlideShape("autoshape-grouping.pptx", 1, "TextBox 7", "SimSun")]
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
    
    [Test]
    [SlideShape("001.pptx", 1, "TextBox 3")]
    [SlideShape("001.pptx", 3, "Text Placeholder 3")]
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
    
    [Test]
    [MasterShape("001.pptx", "Freeform: Shape 7", 18)]
    [SlideShape("020.pptx", 1, 3, 18)]
    [SlideShape("015.pptx", 2, 61, 18.67)]
    [SlideShape("009_table.pptx", 3, 2, 18)]
    [SlideShape("009_table.pptx", 4, 2, 44)]
    [SlideShape("009_table.pptx", 4, 3, 32)]
    [SlideShape("019.pptx", 1, 4103, 18)]
    [SlideShape("019.pptx", 1, 2, 12)]
    [SlideShape("014.pptx", 2, 5, 21.77)]
    [SlideShape("012_title-placeholder.pptx", 1, "Title 1", 20)]
    [SlideShape("010.pptx", 1, 2, 15.39)]
    [SlideShape("014.pptx", 4, 5, 12)]
    [SlideShape("014.pptx", 5, 4, 12)]
    [SlideShape("014.pptx", 6, 52, 27)]
    [SlideShape("autoshape-case016.pptx", 1, "Text Placeholder 1", 28)]
    public void Size_Getter_returns_font_size(IShape shape, double expectedSize)
    {
        // Arrange
        var autoShape =  shape;
        var font = autoShape.TextFrame!.Paragraphs[0].Portions[0].Font;
        
        // Act
        var fontSize = font.Size;
        
        // Assert
        fontSize.Should().Be((decimal)expectedSize);
    }
    
    [Test]
    [SlideShape("028.pptx", 1, 4098, 32)]
    [SlideShape("029.pptx", 1, "Content Placeholder 2", 25)]
    public void Size_Getter_returns_font_size_of_Placeholder(IShape shape, int expectedFontSize)
    {
        // Arrange
        var font = shape.TextFrame!.Paragraphs[0].Portions[0].Font;

        // Act
        var fontSize = font.Size;

        // Assert
        fontSize.Should().Be(expectedFontSize);
    }
    
    [Test]
    [TestCase("001.pptx", 1, "TextBox 3")]
    [TestCase("026.pptx", 1, "AutoShape 1")]
    [TestCase("autoshape-case016.pptx", 1, "Text Placeholder 1")]
    public void Size_Setter_sets_font_size(string presentation, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(StreamOf(presentation));
        var font = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName).TextFrame!.Paragraphs[0].Portions[0].Font;
        var mStream = new MemoryStream();
        var oldSize = font.Size;
        var newSize = oldSize + 2.4m;

        // Act
        font.Size = newSize;

        // Assert
        pres.Validate();
        pres.SaveAs(mStream);
        pres = new Presentation(mStream);
        font = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName).TextFrame!.Paragraphs[0].Portions[0].Font;
        font.Size.Should().Be(newSize);
    }
    
    [Test]
    [TestCase("020.pptx", 3, 7)]
    [TestCase("026.pptx", 1, 128)]
    public void IsBold_Setter_sets_the_placeholder_font_to_be_bold(string presentation, int slideNumber, int shapeId)
    {
        // Arrange
        var pres = new Presentation(StreamOf(presentation));
        var placeholder = pres.Slides[slideNumber - 1].Shapes.GetById<IShape>(shapeId);
        var font = placeholder.TextFrame.Paragraphs[0].Portions[0].Font;
        var mStream = new MemoryStream();

        // Act
        font.IsBold = true;

        // Assert
        font.IsBold.Should().BeTrue();
        pres.SaveAs(mStream);
        pres = new Presentation(mStream);
        placeholder = pres.Slides[slideNumber - 1].Shapes.GetById<IShape>(shapeId);
        font = placeholder.TextFrame.Paragraphs[0].Portions[0].Font;
        font.IsBold.Should().BeTrue();
    }

    [Test]
    [TestCase("autoshape-case010.pptx", 1, "Title 1", 50)]
    [TestCase("autoshape-case010.pptx", 2, "Title 1", -32)]
    public void OffsetEffect_Getter_returns_offset_of_Text(string presentation, int slideNumber, string shapeName, int expectedOffset)
    {
        // Arrange
        var pres = new Presentation(StreamOf(presentation));
        var shape = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName);
        var font = shape.TextFrame!.Paragraphs[0].Portions[1].Font;

        // Act
        var offsetSize = font.OffsetEffect;

        // Assert
        offsetSize.Should().Be(expectedOffset);
    }

    [Test]
    [TestCase("autoshape-case010.pptx", 3, "Title 1", 12)]
    [TestCase("autoshape-case010.pptx", 2, "Title 1", -27)]
    public void OffsetEffect_Setter_changes_Offset_of_paragraph_portion(string presentation, int slideNumber, string shapeName, int expectedOffsetEffect)
    {
        // Arrange
        var pres = new Presentation(StreamOf(presentation));
        var font = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName).TextFrame!.Paragraphs[0].Portions[0].Font;
        var mStream = new MemoryStream();
        var oldOffsetSize = font.OffsetEffect;

        // Act
        font.OffsetEffect = expectedOffsetEffect;
        pres.SaveAs(mStream);

        // Assert
        pres = new Presentation(mStream);
        font = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName).TextFrame!.Paragraphs[0].Portions[0].Font;
        font.OffsetEffect.Should().NotBe(oldOffsetSize);
        font.OffsetEffect.Should().Be(expectedOffsetEffect);
    }
}