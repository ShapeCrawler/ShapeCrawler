using FluentAssertions;
using NUnit.Framework;

namespace ShapeCrawler.Tests.Unit;

public class ParagraphCollectionTests
{
#if DEBUG
    [Test]
    [Explicit("A new feature that should be implemented")]
    public void Add_adds_paragraph()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slide(1);
        slide.Shapes.AddShape(100, 100, 200, 200);
        var addedShape = slide.Shapes.Last();
        var paragraphs = addedShape.TextBox!.Paragraphs;
        paragraphs.Add();
        paragraphs.Last().Text = "Paragraph 1";
        paragraphs.Add();
        paragraphs.Last().Text = "Paragraph 2";
        
        // Act
        paragraphs.Add("New Paragraph 2", 1);
        
        // Assert
        addedShape.TextBox.Text.Should().Be("Paragraph 1\nNew Paragraph 2\nParagraph 2");
    }
#endif
}