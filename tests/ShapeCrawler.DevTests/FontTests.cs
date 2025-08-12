using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;
using ShapeCrawler.Texts;

// ReSharper disable SuggestVarOrType_SimpleTypes

namespace ShapeCrawler.DevTests;

public class FontTests : SCTest
{
    [Test]
    public void EastAsianName_Setter_sets_font_for_the_east_asian_characters_1()
    {
        // Arrange
        var pres = new Presentation(pres =>
        {
            pres.Slide(slide =>
            {
                slide.TextBox(
                    name: "TextBox",
                    x: 100,
                    y: 100,
                    width: 200,
                    height: 50,
                    content: "Test");
            });
        });
        var font = pres.Slide(1).Shapes.Last().TextBox!.Paragraphs[0].Portions[0].Font!;

        // Act
        font.EastAsianName = "SimSun";

        // Assert
        font.EastAsianName.Should().Be("SimSun");
        pres.Validate();
    }
    
    [Test]
    [TestCase("001.pptx", 1, "TextBox 3")]
    public void EastAsianName_Setter_sets_font_for_the_east_asian_characters_2(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(file));
        var shape = pres.Slide(slideNumber).Shapes.Shape(shapeName);
        var font = shape.TextBox!.Paragraphs[0].Portions[0].Font!;

        // Act
        font.EastAsianName = "SimSun";

        // Assert
        font.EastAsianName.Should().Be("SimSun");
        pres.Validate();
    }
    
    [Test]
    [SlideShape("autoshape-grouping.pptx", 1, "TextBox 7", "SimSun")]
    public void EastAsianName_Getter_returns_font_for_East_Asian_characters(IShape shape, string expectedFontName)
    {
        // Arrange
        var font = shape.TextBox!.Paragraphs[0].Portions[0].Font!;

        // Act & Assert
        font.EastAsianName.Should().Be(expectedFontName);
    }
    
    [Test]
    public void Size_Getter_returns_font_size_of_non_first_portion()
    {
        // Arrange
        var pres1 = new Presentation(TestAsset("015.pptx"));
        var font1 = pres1.Slide(1).Shapes.GetById<IShape>(5).TextBox!.Paragraphs[0].Portions[2].Font!;
        var pres2 = new Presentation(TestAsset("009_table.pptx"));
        var font2 = pres2.Slide(3).Shapes.GetById<IShape>(2).TextBox!.Paragraphs[0].Portions[1].Font!;

        // Act
        var fontSize1 = font1.Size;
        var fontSize2 = font2.Size;
        
        // Assert
        fontSize1.Should().Be(18);
        fontSize2.Should().Be(20);
    }

    [Test]
    public void Size_Getter_returns_font_size_of_Non_Placeholder_Table()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));
        var table = pres.Slide(3).Shape(3);
        var portion = table.Table.Rows[0].Cells[0].TextBox.Paragraphs[0].Portions[0];

        // Act-Assert
        portion.Font!.Size.Should().Be(18);
    }
    
    [Test]
    [MasterShape("001.pptx", "Freeform: Shape 7", 18)]
    [SlideShape("020.pptx", 1, 3, 18)]
    [SlideShape("015.pptx", 2, 61, 18.67)]
    [SlideShape("009_table.pptx", 3, 2, 18)]
    [SlideShape("009_table.pptx", 4, 2, 44)]
    [SlideShape("009_table.pptx", 4, 3, 32)]
    [SlideShape("019.pptx", 1, 4103, 18)]
    [SlideShape("019.pptx", 1, "Slide Number", 12)]
    [SlideShape("014.pptx", 2, 5, 21.77)]
    [SlideShape("012_title-placeholder.pptx", 1, "Title 1", 20)]
    [SlideShape("010.pptx", 1, 2, 15.39)]
    [SlideShape("014.pptx", 4, 5, 12)]
    [SlideShape("014.pptx", 5, 4, 12)]
    [SlideShape("014.pptx", 6, 52, 27)]
    [SlideShape("autoshape-case016.pptx", 1, "Text Placeholder 1", 28)]
    [SlideShape("001.pptx", 1, "TextBox 8", 11)]
    public void Size_Getter_returns_font_size(IShape shape, double expectedSize)
    { 
        // Arrange
        var font = shape.TextBox!.Paragraphs[0].Portions[0].Font!;
        
        // Act & Assert
        font.Size.Should().Be((decimal)expectedSize);
    }
    
    [Test]
    [SlideShape("028.pptx", 1, 4098, 32)]
    [SlideShape("029.pptx", 1, "Content Placeholder 2", 25)]
    [SlideShape("072 content placeholder.pptx", 1, "Content Placeholder 1", 18)]
    public void Size_Getter_returns_font_size_of_Placeholder(IShape shape, int expectedSize)
    {
        // Arrange
        var font = shape.TextBox!.Paragraphs[0].Portions[0].Font;

        // Act
        var fontSize = font!.Size;

        // Assert
        fontSize.Should().Be(expectedSize);
    }

    [Test]
    public void IsBold_GetterReturnsTrue_WhenFontOfNonPlaceholderTextIsBold()
    {
        // Arrange
        var pres = new Presentation(TestAsset("020.pptx"));
        var nonPlaceholder = pres.Slide(1).Shapes.GetById(3);
        var font = nonPlaceholder.TextBox!.Paragraphs[0].Portions[0].Font!;

        // Act-Assert
        font.IsBold.Should().BeTrue();
    }

    [Test]
    public void IsBold_GetterReturnsTrue_WhenFontOfPlaceholderTextIsBold()
    {
        // Arrange
        var pres = new Presentation(TestAsset("020.pptx"));
        var placeholder = pres.Slide(2).Shapes.GetById<IShape>(6);
        var font = placeholder.TextBox!.Paragraphs[0].Portions[0].Font!;

        // Act & Assert
        font.IsBold.Should().BeTrue();
    }

    [Test]
    public void IsBold_GetterReturnsFalse_WhenFontOfNonPlaceholderTextIsNotBold()
    {
        // Arrange
        var pres = new Presentation(TestAsset("020.pptx"));
        var nonPlaceholder = pres.Slide(1).Shapes.GetById(2);
        var font = nonPlaceholder.TextBox!.Paragraphs[0].Portions[0].Font!;

        // Act & Assert
        font.IsBold.Should().BeFalse();
    }

    [Test]
    public void IsBold_GetterReturnsFalse_WhenFontOfPlaceholderTextIsNotBold()
    {
        // Arrange
        var pres = new Presentation(TestAsset("020.pptx"));
        var placeholder = pres.Slide(3).Shapes.First(sp => sp.Id == 7);
        var font = placeholder.TextBox!.Paragraphs[0].Portions[0].Font!;
        
        // Act & Assert
        font.IsBold.Should().BeFalse();
    }

    [Test]
    public void IsBold_Setter_AddsBoldForNonPlaceholderTextFont()
    {
        // Arrange
        var mStream = new MemoryStream();
        var pres = new Presentation(TestAsset("020.pptx"));
        var font = pres.Slide(1).Shapes.GetById(2).TextBox!.Paragraphs[0].Portions[0].Font!;

        // Act
        font.IsBold = true;

        // Assert
        font.IsBold.Should().BeTrue();
        pres.Save(mStream);
        pres = new Presentation(mStream);
        font = pres.Slide(1).Shapes.GetById(2).TextBox!.Paragraphs[0].Portions[0].Font!;
        font.IsBold.Should().BeTrue();
    }

    [Test]
    public void IsItalic_GetterReturnsTrue_WhenFontOfNonPlaceholderTextIsItalic()
    {
        // Arrange
        var pres = new Presentation(TestAsset("020.pptx"));
        var nonPlaceholder = pres.Slide(1).Shapes.GetById(3);
        var font = nonPlaceholder.TextBox!.Paragraphs[0].Portions[0].Font!;

        // Act & Assert
        font.IsItalic.Should().BeTrue();
    }

    [Test]
    public void IsItalic_GetterReturnsTrue_WhenFontOfPlaceholderTextIsItalic()
    {
        // Arrange
        var pres = new Presentation(TestAsset("020.pptx"));
        var placeholder = pres.Slide(3).Shapes.GetById(7);
        var font = placeholder.TextBox!.Paragraphs[0].Portions[0].Font!;

        // Act-Assert
        font.IsItalic.Should().BeTrue();
    }

    [Test]
    public void IsItalic_Setter_SetsItalicFontForForNonPlaceholderText()
    {
        // Arrange
        var mStream = new MemoryStream();
        var pres = new Presentation(TestAsset("020.pptx"));
        var nonPlaceholder = pres.Slide(1).Shapes.GetById(2);
        var font = nonPlaceholder.TextBox!.Paragraphs[0].Portions[0].Font!;

        // Act
        font.IsItalic = true;

        // Assert
        font.IsItalic.Should().BeTrue();
        pres.Save(mStream);
        pres = new Presentation(mStream);
        nonPlaceholder = pres.Slide(1).Shapes.GetById(2);
        font = nonPlaceholder.TextBox!.Paragraphs[0].Portions[0].Font!;
        font.IsItalic.Should().BeTrue();
    }

    [Test]
    public void IsItalic_SetterSetsNonItalicFontForPlaceholderText_WhenFalseValueIsPassed()
    {
        // Arrange
        var mStream = new MemoryStream();
        var pres = new Presentation(TestAsset("020.pptx"));
        var placeholder = pres.Slide(3).Shapes.GetById(7);
        var font = placeholder.TextBox!.Paragraphs[0].Portions[0].Font!;

        // Act
        font.IsItalic = false;

        // Assert
        font.IsItalic.Should().BeFalse();
        pres.Save(mStream);
        pres = new Presentation(mStream);
        placeholder = pres.Slide(3).Shapes.GetById(7);
        font = placeholder.TextBox!.Paragraphs[0].Portions[0].Font!;
        font.IsItalic.Should().BeFalse();
    }

    [Test]
    public void Underline_SetUnderlineFont_WhenValueEqualsSetPassed()
    {
        // Arrange
        var mStream = new MemoryStream();
        var pres = new Presentation(TestAsset("020.pptx"));
        var placeholder = pres.Slide(3).Shapes.GetById(7);
        var font = placeholder.TextBox!.Paragraphs[0].Portions[0].Font!;

        // Act
        font.Underline = DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Single;

        // Assert
        font.Underline.Should().Be(DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Single);
        pres.Save(mStream);
        pres = new Presentation(mStream);
        placeholder = pres.Slide(3).Shapes.GetById(7);
        font = placeholder.TextBox!.Paragraphs[0].Portions[0].Font!;
        font.Underline.Should().Be(DocumentFormat.OpenXml.Drawing.TextUnderlineValues.Single);
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
        var font = shape.TextBox!.Paragraphs[0].Portions[0].Font!;

        // Act & Assert
        font.LatinName.Should().Be(expectedFontName);
    }
    
    [Test]
    [SlideShape("001.pptx", 1, "TextBox 3")]
    [SlideShape("001.pptx", 3, "Text Placeholder 3")]
    public void LatinName_Setter_sets_font_for_the_latin_characters(IShape shape)
    {
        // Arrange
        var font = shape.TextBox!.Paragraphs[0].Portions[0].Font!;

        // Act
        font.LatinName = "Time New Roman";

        // Assert
        font.LatinName.Should().Be("Time New Roman");
    }
    
    [Test]
    public void LatinName_Setter_sets_font_for_the_latin_characters_of_table_cell()
    {
        // Arrange
        var pres = new Presentation(pres =>
        {
            pres.Slide(slide =>
            {
                slide.Table("Table 1", x: 40, y: 40, columnsCount: 6, rowsCount: 5);
            });
        });
        var slide = pres.Slide(1);
        var table = (ITable)slide.Shapes.Last().Table;
        var cell = table[1, 2];
        cell.TextBox.SetText("Test");
        var font = cell.TextBox.Paragraphs.First().Portions.First().Font!;

        // Act
        font.LatinName = "Arial";
        
        // Assert
        font.LatinName.Should().Be("Arial");
    }
    
    [Test]
    public void LatinName_Setter_sets_font_for_the_latin_characters()
    {
        // Arrange
        var pres = new Presentation(pres =>
        {
            pres.Slide(slide =>
            {
                slide.TextBox("TextBox 1", x: 10, y: 20, width: 30, height: 40, content: "Shape 1");
                slide.TextBox("TextBox 2", x: 100, y: 20, width: 30, height: 40, content: "Test");
            });
        });
        var shape1 = pres.Slide(1).Shape("TextBox 1");
        var shape2 = pres.Slide(1).Shape("TextBox 2");
        shape1.TextBox!.Paragraphs[0].Portions[0].Font!.LatinName = "Segoe UI Semibold";

        // Act
        shape2.TextBox!.SetText("Shape 2");
        shape2.TextBox!.Paragraphs[0].Portions[0].Font!.LatinName = "Aptos";

        // Assert
        shape1.TextBox!.Paragraphs[0].Portions[0].Font.LatinName.Should().Be("Segoe UI Semibold");
    }
    
    [Test]
    [TestCase("001.pptx", 1, "TextBox 3")]
    [TestCase("026.pptx", 1, "AutoShape 1")]
    [TestCase("autoshape-case016.pptx", 1, "Text Placeholder 1")]
    public void Size_Setter_sets_font_size(string presentation, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(presentation));
        var font = pres.Slide(slideNumber).Shape(shapeName).TextBox!.Paragraphs[0].Portions[0].Font!;
        var mStream = new MemoryStream();
        var oldSize = font.Size;
        var newSize = oldSize + 2.4m;

        // Act
        font.Size = newSize;

        // Assert
        pres.Validate();
        pres.Save(mStream);
        pres = new Presentation(mStream);
        font = pres.Slide(slideNumber).Shape(shapeName).TextBox!.Paragraphs[0].Portions[0].Font!;
        font.Size.Should().Be(newSize);
    }
    
    [Test]
    [TestCase("020.pptx", 3, 7)]
    [TestCase("026.pptx", 1, 128)]
    public void IsBold_Setter_sets_the_placeholder_font_to_be_bold(string presentation, int slideNumber, int shapeId)
    {
        // Arrange
        var pres = new Presentation(TestAsset(presentation));
        var placeholder = pres.Slide(slideNumber).Shape(shapeId);
        var font = placeholder.TextBox!.Paragraphs[0].Portions[0].Font!;
        var mStream = new MemoryStream();

        // Act
        font.IsBold = true;

        // Assert
        font.IsBold.Should().BeTrue();
        pres.Save(mStream);
        pres = new Presentation(mStream);
        placeholder = pres.Slide(slideNumber).Shape(shapeId);
        font = placeholder.TextBox!.Paragraphs[0].Portions[0].Font!;
        font.IsBold.Should().BeTrue();
    }

    [Test]
    [TestCase("autoshape-case010.pptx", 1, "Title 1", 50)]
    [TestCase("autoshape-case010.pptx", 2, "Title 1", -32)]
    public void OffsetEffect_Getter_returns_offset_of_Text(string presentation, int slideNumber, string shapeName, int expectedOffset)
    {
        // Arrange
        var pres = new Presentation(TestAsset(presentation));
        var shape = pres.Slide(slideNumber).Shape(shapeName);
        var font = shape.TextBox!.Paragraphs[0].Portions[1].Font!;

        // Act & Assert
        font.OffsetEffect.Should().Be(expectedOffset);
    }

    [Test]
    [TestCase("autoshape-case010.pptx", 3, "Title 1", 12)]
    [TestCase("autoshape-case010.pptx", 2, "Title 1", -27)]
    public void OffsetEffect_Setter_changes_Offset_of_paragraph_portion(string presentation, int slideNumber, string shapeName, int expectedOffsetEffect)
    {
        // Arrange
        var pres = new Presentation(TestAsset(presentation));
        var font = pres.Slide(slideNumber).Shape(shapeName).TextBox!.Paragraphs[0].Portions[0].Font!;
        var mStream = new MemoryStream();
        var oldOffsetSize = font.OffsetEffect;

        // Act
        font.OffsetEffect = expectedOffsetEffect;
        pres.Save(mStream);

        // Assert
        pres = new Presentation(mStream);
        font = pres.Slide(slideNumber).Shape(shapeName).TextBox!.Paragraphs[0].Portions[0].Font!;
        font.OffsetEffect.Should().NotBe(oldOffsetSize);
        font.OffsetEffect.Should().Be(expectedOffsetEffect);
    }
    
    [Test]
    public void LatinName_Setter()
    {
        // Arrange
        var pres = new Presentation(pres =>
        {
            pres.Slide(slide =>
            {
                slide.TextBox("TextBox 1", 0, 0, 100, 100, "Test");
            });
        });
        var slide = pres.Slide(1);
        var addedShape = slide.Shapes.Last();
        var font = addedShape.TextBox!.Paragraphs[0].Portions[0].Font!;
        var stream = new MemoryStream();
        
        // Act
        font.LatinName = "Times New Roman";
        
        // Assert
        pres.Save(stream);
        font = new Presentation(stream).Slide(1).Shapes.Last().TextBox!.Paragraphs[0].Portions[0].Font!;
        font.LatinName.Should().Be("Times New Roman");
        pres.Validate();
    }
}
