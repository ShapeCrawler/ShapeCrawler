using FluentAssertions;
using NUnit.Framework;

namespace ShapeCrawler.DevTests;

public class ParagraphCollectionTests
{
    [Test]
    public void Add_adds_paragraph_in_the_middle()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slide(1);
        slide.Shapes.AddShape(100, 100, 200, 200);
        var addedShape = slide.Shapes.Last();
        var paragraphs = addedShape.TextBox!.Paragraphs;
        paragraphs[0].Text = "Paragraph 1";
        paragraphs.Add();
        paragraphs.Last().Text = "Paragraph 2";

        // Act
        paragraphs.Add("New Paragraph 2", 1);

        // Assert
        addedShape.TextBox.Text.Should().Be($"Paragraph 1{Environment.NewLine}New Paragraph 2{Environment.NewLine}Paragraph 2");
        pres.Validate();
    }
    
    [Test]
    public void Add_adds_paragraph_at_the_beginning()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slide(1);
        slide.Shapes.AddShape(100, 100, 200, 200);
        var addedShape = slide.Shapes.Last();
        var paragraphs = addedShape.TextBox!.Paragraphs;
        paragraphs[0].Text = "Paragraph 1";

        // Act
        paragraphs.Add("New Paragraph 1", 0);

        // Assert
        addedShape.TextBox.Text.Should().Be($"New Paragraph 1{Environment.NewLine}Paragraph 1");
        pres.Validate();
    }
}