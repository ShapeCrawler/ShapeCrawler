using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tests;

public partial class UserSlideShapeCollectionTests
{
    [Test]
    public void AddMarkdownTextBox_preserves_inline_formatting_and_editable_text()
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());
        var shapes = presentation.Slide(1).Shapes;

        // Act
        shapes.AddMarkdownTextBox(50, 60, 400, 200, "Plain **bold** *italic* ***both*** [**link**](https://example.com)");

        // Assert
        using var reopened = SaveAndOpenPresentation(presentation);
        var shape = reopened.Slide(1).Shapes.Single();
        shape.X.Should().Be(50);
        shape.Y.Should().Be(60);
        shape.Width.Should().Be(400);
        shape.Height.Should().Be(200);
        shape.TextBox.Text.Should().Be("Plain bold italic both link");
        using var sdk = SaveAndOpenPresentationAsSdk(reopened);
        var slide = sdk.PresentationPart!.SlideParts.Single();
        var runs = slide.Slide.Descendants<A.Run>().ToArray();
        runs.Single(r => r.InnerText == "bold").RunProperties!.Bold!.Value.Should().BeTrue();
        runs.Single(r => r.InnerText == "italic").RunProperties!.Italic!.Value.Should().BeTrue();
        var both = runs.Single(r => r.InnerText == "both").RunProperties!;
        both.Bold!.Value.Should().BeTrue();
        both.Italic!.Value.Should().BeTrue();
        runs[0].RunProperties!.Bold!.Value.Should().BeFalse();
        var link = runs.Single(r => r.InnerText == "link").RunProperties!;
        link.Bold!.Value.Should().BeTrue();
        slide.HyperlinkRelationships.Single(r => r.Id == link.GetFirstChild<A.HyperlinkOnClick>()!.Id)
            .Uri.ToString().Should().Be("https://example.com/");
        shape.TextBox.Paragraphs[0].Portions[0].Text = "Edited ";
        using var edited = SaveAndOpenPresentation(reopened);
        edited.Slide(1).Shapes.Single().TextBox.Text.Should().StartWith("Edited bold");
        ValidatePresentation(reopened);
    }

    [Test]
    public void AddMarkdownTextBox_preserves_paragraphs_and_nested_lists()
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());
        const string markdown = "# Heading\n\nFirst paragraph\n\nSecond paragraph\n\n- One\n  - Nested\n- Two\n\n3. Third\n4. Fourth";

        // Act
        presentation.Slide(1).Shapes.AddMarkdownTextBox(0, 0, 500, 400, markdown);

        // Assert
        using var sdk = SaveAndOpenPresentationAsSdk(presentation);
        var paragraphs = sdk.PresentationPart!.SlideParts.Single().Slide.Descendants<P.Shape>().Single()
            .TextBody!.Elements<A.Paragraph>().ToArray();
        paragraphs.Select(p => p.InnerText).Should().Equal("Heading", "First paragraph", "Second paragraph",
            "One", "Nested", "Two", "Third", "Fourth");
        paragraphs[0].GetFirstChild<A.Run>()!.RunProperties!.Bold!.Value.Should().BeTrue();
        paragraphs.Take(3).Should().OnlyContain(p => p.ParagraphProperties!.GetFirstChild<A.NoBullet>() != null);
        paragraphs[3].ParagraphProperties!.GetFirstChild<A.CharacterBullet>()!.Char!.Value.Should().Be("•");
        paragraphs[4].ParagraphProperties!.Level!.Value.Should().Be(1);
        paragraphs[5].ParagraphProperties!.Level!.Value.Should().Be(0);
        paragraphs[6].ParagraphProperties!.GetFirstChild<A.AutoNumberedBullet>()!.StartAt!.Value.Should().Be(3);
        paragraphs[7].ParagraphProperties!.GetFirstChild<A.AutoNumberedBullet>()!.StartAt!.Value.Should().Be(4);
        ValidatePresentation(presentation);
    }

    [Test]
    public void AddMarkdownTextBox_preserves_list_continuation_paragraphs()
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());

        // Act
        presentation.Slide(1).Shapes.AddMarkdownTextBox(0, 0, 500, 400, "- First\n\n  Continuation\n\n- Second");

        // Assert
        using var sdk = SaveAndOpenPresentationAsSdk(presentation);
        var paragraphs = sdk.PresentationPart!.SlideParts.Single().Slide.Descendants<A.Paragraph>().ToArray();
        paragraphs.Select(p => p.InnerText).Should().Equal("First", "Continuation", "Second");
        paragraphs[1].ParagraphProperties!.GetFirstChild<A.NoBullet>().Should().NotBeNull();
        paragraphs[1].ParagraphProperties!.LeftMargin!.Value.Should().Be(285750);
        ValidatePresentation(presentation);
    }

    [TestCase("line\nwrapped  \nbreak", "line wrappedbreak", 1)]
    [TestCase("`**code**` &amp; \\*literal\\*", "**code** & *literal*", 0)]
    [TestCase("<b>HTML</b>", "<b>HTML</b>", 0)]
    [TestCase("<div>block</div>", "<div>block</div>", 0)]
    [TestCase("![alt text](missing.png)", "alt text", 0)]
    [TestCase("~~unsupported~~", "~~unsupported~~", 0)]
    [TestCase("> **quoted**", "quoted", 0)]
    [TestCase("before\n\n---\n\nafter", "beforeafter", 0)]
    [TestCase("```csharp\nvar x = 1;\nvar y = 2;\n```", "var x = 1;var y = 2;", 1)]
    public void AddMarkdownTextBox_handles_breaks_and_fallbacks(string markdown, string text, int breaks)
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());

        // Act
        presentation.Slide(1).Shapes.AddMarkdownTextBox(0, 0, 400, 200, markdown);

        // Assert
        using var sdk = SaveAndOpenPresentationAsSdk(presentation);
        var body = sdk.PresentationPart!.SlideParts.Single().Slide.Descendants<P.Shape>().Single().TextBody!;
        body.InnerText.Should().Be(text);
        body.Descendants<A.Break>().Should().HaveCount(breaks);
        ValidatePresentation(presentation);
    }

    [TestCase("[reference][id]\n\n[id]: https://example.com", "https://example.com/")]
    [TestCase("<https://example.com>", "https://example.com/")]
    [TestCase("<a@example.com>", "mailto:a@example.com")]
    [TestCase("[`code`](https://example.com)", "https://example.com/")]
    public void AddMarkdownTextBox_preserves_link_relationships(string markdown, string target)
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());

        // Act
        presentation.Slide(1).Shapes.AddMarkdownTextBox(0, 0, 400, 200, markdown);

        // Assert
        using var sdk = SaveAndOpenPresentationAsSdk(presentation);
        var slide = sdk.PresentationPart!.SlideParts.Single();
        slide.HyperlinkRelationships.Single().Uri.ToString().Should().Be(target);
        slide.Slide.Descendants<A.HyperlinkOnClick>().Single().Id!.Value
            .Should().Be(slide.HyperlinkRelationships.Single().Id);
        ValidatePresentation(presentation);
    }

    [TestCase("")]
    [TestCase("  \n\n")]
    public void AddMarkdownTextBox_creates_valid_empty_text_box(string markdown)
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());

        // Act
        presentation.Slide(1).Shapes.AddMarkdownTextBox(0, 0, 400, 200, markdown);

        // Assert
        using var reopened = SaveAndOpenPresentation(presentation);
        reopened.Slide(1).Shapes.Single().TextBox.Text.Should().BeEmpty();
        reopened.Slide(1).Shapes.Single().TextBox.Paragraphs.Should().HaveCount(1);
        ValidatePresentation(reopened);
    }

    [Test]
    public void AddMarkdownTextBox_rejects_null_without_adding_a_shape()
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());
        var shapes = presentation.Slide(1).Shapes;

        // Act
        var action = () => shapes.AddMarkdownTextBox(0, 0, 400, 200, null!);

        // Assert
        action.Should().Throw<ArgumentNullException>().WithParameterName("markdown");
        shapes.Should().BeEmpty();
    }

    [Test]
    public void AddMarkdownTextBox_caps_deep_list_indentation_and_preserves_other_shapes()
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());
        var shapes = presentation.Slide(1).Shapes;
        shapes.AddTextBox(10, 10, 100, 50, "Existing");
        var markdown = string.Join("\n", Enumerable.Range(0, 11).Select(i => new string(' ', i * 2) + "- Item"));

        // Act
        shapes.AddMarkdownTextBox(0, 0, 500, 400, markdown);
        shapes.AddMarkdownTextBox(0, 400, 500, 100, "**Another**");

        // Assert
        using var reopened = SaveAndOpenPresentation(presentation);
        reopened.Slide(1).Shapes.Should().HaveCount(3);
        reopened.Slide(1).Shapes[0].TextBox.Text.Should().Be("Existing");
        reopened.Slide(1).Shapes[2].TextBox.Text.Should().Be("Another");
        using var sdk = SaveAndOpenPresentationAsSdk(reopened);
        var slide = sdk.PresentationPart!.SlideParts.Single();
        var paragraphs = slide.Slide.Descendants<P.Shape>().ElementAt(1).TextBody!.Elements<A.Paragraph>();
        paragraphs.Select(p => p.ParagraphProperties!.Level!.Value).Should().Equal(0, 1, 2, 3, 4, 5, 6, 7, 8, 8, 8);
        slide.Slide.Descendants<P.Shape>().Select(p => p.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id!.Value)
            .Should().OnlyHaveUniqueItems();
        ValidatePresentation(reopened);
    }
}
