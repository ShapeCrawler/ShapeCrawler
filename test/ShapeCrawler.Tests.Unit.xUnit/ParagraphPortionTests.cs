using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Drawing;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit;

[SuppressMessage("Usage", "xUnit1013:Public method should be marked as test")]
public class ParagraphPortionTests : SCTest
{
    [Fact]
    public void Text_Getter_returns_text_of_paragraph_portion()
    {
        // Arrange
        var pptx = StreamOf("009_table");
        var pres = new SCPresentation(pptx);
        IParagraphPortion portion = ((ITable)pres.Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0]
            .TextFrame
            .Paragraphs[0].Portions[0];

        // Act
        string paragraphPortionText = portion.Text;

        // Assert
        paragraphPortionText.Should().BeEquivalentTo("0:0_p1_lvl1");
    }

    [Fact]
    public void Text_Setter_updates_text()
    {
        // Arrange
        var pptxStream = StreamOf("autoshape-case001.pptx");
        var pres = new SCPresentation(pptxStream);
        var autoShape = pres.SlideMasters[0].Shapes.GetByName<IShape>("AutoShape 1");
        var portion = autoShape.TextFrame!.Paragraphs[0].Portions[0];

        // Act
        portion.Text = "test";

        // Assert
        portion.Text.Should().Be("test");
    }

    [Xunit.Theory]
    [MemberData(nameof(TestCasesHyperlinkSetter))]
    public void Hyperlink_Setter_sets_hyperlink(string pptxFile, string shapeName)
    {
        // Arrange
        var pptxStream = StreamOf(pptxFile);
        var presentation = new SCPresentation(pptxStream);
        var autoShape = presentation.Slides[0].Shapes.GetByName<IShape>(shapeName);
        var portion = autoShape.TextFrame.Paragraphs[0].Portions[0];

        // Act
        portion.Hyperlink = "https://github.com/ShapeCrawler/ShapeCrawler";

        // Assert
        presentation.Save();
        presentation = new SCPresentation(pptxStream);
        autoShape = presentation.Slides[0].Shapes.GetByName<IShape>(shapeName);
        portion = autoShape.TextFrame.Paragraphs[0].Portions[0];
        portion.Hyperlink.Should().Be("https://github.com/ShapeCrawler/ShapeCrawler");
    }

    public static IEnumerable<object[]> TestCasesHyperlinkSetter()
    {
        yield return new[] { "001.pptx", "TextBox 3" };
        yield return new[] { "autoshape-case001.pptx", "AutoShape 1" };
        yield return new[] { "autoshape-case002.pptx", "AutoShape 1" };
    }

    [Fact]
    public void Hyperlink_Setter_sets_hyperlink_for_two_shape_on_the_Same_slide()
    {
        // Arrange
        var pptxStream = StreamOf("001.pptx");
        var presentation = new SCPresentation(pptxStream);
        var textBox3 = presentation.Slides[0].Shapes.GetByName<IShape>("TextBox 3");
        var textBox4 = presentation.Slides[0].Shapes.GetByName<IShape>("TextBox 4");
        var portion3 = textBox3.TextFrame.Paragraphs[0].Portions[0];
        var portion4 = textBox4.TextFrame.Paragraphs[0].Portions[0];

        // Act
        portion3.Hyperlink = "https://github.com/ShapeCrawler/ShapeCrawler";
        portion4.Hyperlink = "https://github.com/ShapeCrawler/ShapeCrawler";

        // Assert
        portion3.Hyperlink.Should().Be("https://github.com/ShapeCrawler/ShapeCrawler");
        portion4.Hyperlink.Should().Be("https://github.com/ShapeCrawler/ShapeCrawler");
    }

    [Fact]
    public void Hyperlink_Setter_sets_File_Name_as_a_hyperlink()
    {
        // Arrange
        var pptx = StreamOf("autoshape-case001.pptx");
        var pres = new SCPresentation(pptx);
        var shape = pres.Slides[0].Shapes.GetByName<IShape>("AutoShape 1");
        var portion = shape.TextFrame!.Paragraphs[0].Portions[0];
        
        // Act
        portion.Hyperlink = "some.pptx";
        
        // Assert
        portion.Hyperlink.Should().Be("some.pptx");
    }
    
    [Fact]
    public void Hyperlink_Setter_sets_hyperlink_for_table_Cell()
    {
        // Arrange
        var pptxStream =  StreamOf("table-case001.pptx");
        var pres = new SCPresentation(pptxStream);
        var table = pres.Slides[0].Shapes.GetByName<ITable>("Table 1");
        var portion = table.Rows[0].Cells[0].TextFrame.Paragraphs[0].Portions[0];

        // Act
        portion.Hyperlink = "https://github.com/ShapeCrawler/ShapeCrawler";

        // Assert
        portion.Hyperlink.Should().Be("https://github.com/ShapeCrawler/ShapeCrawler");
        pres.Validate();
    }

    [Fact]
    public void TextHighlightColor_Getter_returns_text_highlight_color()
    {
        // Arrange
        var pptx = StreamOf("autoshape-grouping.pptx");
        var pres = new SCPresentation(pptx);
        var shape = pres.Slides[0].Shapes.GetByName<IShape>("TextBox 3");
        var portion = shape.TextFrame!.Paragraphs[0].Portions[0];

        // Act-Assert
        portion.TextHighlightColor.ToString().Should().Be("FFFF00");
    }

    [Fact]
    public void TextHighlightColor_Getter_returns_text_highlight_sccolor()
    {
        // Arrange
        var pptx = StreamOf("autoshape-grouping.pptx");
        var pres = new SCPresentation(pptx);
        var shape = pres.Slides[0].Shapes.GetByName<IShape>("TextBox 3");
        var portion = shape.TextFrame!.Paragraphs[0].Portions[0];

        // Act-Assert
        portion.TextHighlightColor.ToString().Should().Be("FFFF00");
    }

    [Fact]
    public void TextHighlightColor_Setter_sets_text_highlight_color()
    {
        // Arrange
        var pptx = StreamOf("autoshape-grouping.pptx");
        var pres = new SCPresentation(pptx);
        var shape = pres.Slides[0].Shapes.GetByName<IShape>("TextBox 4");
        var portion = shape.TextFrame!.Paragraphs[0].Portions[0];

        // Act
        portion.TextHighlightColor = SCColor.FromHex("FFFF00");

        // Assert
        portion.TextHighlightColor.ToString().Should().Be("FFFF00");
    }

    [Fact]
    public void TextHighlightColor_Setter_sets_text_highlight()
    {
        // Arrange
        var pptx = StreamOf("autoshape-grouping.pptx");
        var pres = new SCPresentation(pptx);
        var shape = pres.Slides[0].Shapes.GetByName<IShape>("TextBox 4");
        var portion = shape.TextFrame!.Paragraphs[0].Portions[0];
        var color = SCColor.FromHex("FFFF00");

        // Act
        portion.TextHighlightColor = color;

        // Assert
        portion.TextHighlightColor.ToString().Should().Be("FFFF00");
    }
}