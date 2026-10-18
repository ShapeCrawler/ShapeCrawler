using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using ShapeCrawler.Tests.Helpers;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tests;

public class HyperlinkTests : SCTest
{
    [Test]
    public void Remove_preserves_shared_URL_relationship_until_last_reference_is_removed()
    {
        // Arrange
        using var pres = LinkPresentation();
        var shape = pres.Slides[0].Shapes.Last();
        var textLink = shape.TextBox!.Paragraphs[0].Portions[0].Link!;
        shape.Link.AddFile("https://example.com/report");
        textLink.AddFile("https://example.com/report");

        // Act
        shape.Link.Remove();

        // Assert
        shape.Link.File.Should().BeNull();
        textLink.File.Should().Be("https://example.com/report");
        RelationshipCount(pres).Should().Be(1);

        // Act
        textLink.Remove();

        // Assert
        RelationshipCount(pres).Should().Be(0);
        shape.TextBox.Text.Should().Be("Button");
        ValidatePresentation(pres);
    }

    [Test]
    public void AddFile_replaces_link_without_accumulating_unused_relationships()
    {
        // Arrange
        using var pres = LinkPresentation();
        var link = pres.Slides[0].Shapes.Last().Link;
        link.AddFile("https://example.com/old");

        // Act
        link.AddFile("https://example.com/new");
        link.AddFile("https://example.com/new");

        // Assert
        link.File.Should().Be("https://example.com/new");
        RelationshipCount(pres).Should().Be(1);
        ValidatePresentation(pres);
    }

    [Test]
    public void AddFile_keeps_different_URL_fragments_distinct()
    {
        // Arrange
        using var pres = LinkPresentation();
        var shape = pres.Slides[0].Shapes.Last();
        shape.Link.AddFile("https://example.com/report#one");
        var textLink = shape.TextBox!.Paragraphs[0].Portions[0].Link!;
        textLink.AddFile("https://example.com/report#one");

        // Act
        shape.Link.AddFile("https://example.com/report#two");

        // Assert
        shape.Link.File.Should().Be("https://example.com/report#two");
        textLink.File.Should().Be("https://example.com/report#one");
        RelationshipCount(pres).Should().Be(2);
        ValidatePresentation(pres);
    }

    [Test]
    public void AddFile_replaces_internal_link_without_deleting_target_slide()
    {
        // Arrange
        using var pres = LinkPresentation();
        pres.Slides.Add(1);
        var link = pres.Slides[0].Shapes.Last().Link;
        link.AddSlideNumber(2);

        // Act
        link.AddFile("https://example.com/report");

        // Assert
        link.File.Should().Be("https://example.com/report");
        link.SlideNumber.Should().BeNull();
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var reopened = new Presentation(saved);
        reopened.Slides.Should().HaveCount(2);
        ValidatePresentation(reopened);
    }

    [Test]
    public void AddFile_rejects_malformed_URI_before_changing_existing_link()
    {
        // Arrange
        using var pres = LinkPresentation();
        var link = pres.Slides[0].Shapes.Last().Link;
        link.AddFile("https://example.com/old");

        // Act
        var add = () => link.AddFile("http://[");

        // Assert
        add.Should().Throw<UriFormatException>();
        link.File.Should().Be("https://example.com/old");
        RelationshipCount(pres).Should().Be(1);
    }

    [Test]
    public void AddSlideNumber_replaces_external_link_and_tracks_slide_reordering()
    {
        // Arrange
        using var pres = LinkPresentation();
        pres.Slides.Add(1);
        var link = pres.Slides[0].Shapes.Last().Link;
        link.AddFile("https://example.com/old");

        // Act
        link.AddSlideNumber(2);
        pres.Slides[1].Number = 1;

        // Assert
        link.SlideNumber.Should().Be(1);
        link.File.Should().BeNull();
        RelationshipCount(pres).Should().Be(0);
        ValidatePresentation(pres);
    }

    [Test]
    public void Remove_internal_link_keeps_target_slide_and_is_idempotent()
    {
        // Arrange
        using var pres = LinkPresentation();
        pres.Slides.Add(1);
        var link = pres.Slides[0].Shapes.Last().Link;
        link.AddSlideNumber(2);

        // Act
        link.Remove();
        link.Remove();

        // Assert
        link.SlideNumber.Should().BeNull();
        link.File.Should().BeNull();
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var reopened = new Presentation(saved);
        reopened.Slides.Should().HaveCount(2);
        ValidatePresentation(reopened);
    }

    [TestCase(0)]
    [TestCase(-1)]
    [TestCase(99)]
    public void AddSlideNumber_rejects_invalid_targets_without_changing_existing_link(int target)
    {
        // Arrange
        using var pres = LinkPresentation();
        var link = pres.Slides[0].Shapes.Last().Link;
        link.AddFile("https://example.com/old");

        // Act
        var add = () => link.AddSlideNumber(target);

        // Assert
        add.Should().Throw<ArgumentOutOfRangeException>();
        link.File.Should().Be("https://example.com/old");
        RelationshipCount(pres).Should().Be(1);
    }

    [TestCase(null)]
    [TestCase("")]
    [TestCase(" ")]
    public void AddFile_rejects_empty_addresses_without_creating_link(string? address)
    {
        // Arrange
        using var pres = LinkPresentation();
        var link = pres.Slides[0].Shapes.Last().Link;

        // Act
        var add = () => link.AddFile(address!);

        // Assert
        add.Should().Throw<ArgumentException>();
        link.File.Should().BeNull();
        RelationshipCount(pres).Should().Be(0);
    }

    [Test]
    public void Remove_preserves_mouse_over_link_using_same_relationship()
    {
        // Arrange
        using var source = LinkPresentation();
        source.Slides[0].Shapes.Last().Link.AddFile("https://example.com/hover");
        using var stream = new MemoryStream();
        source.Save(stream);
        using (var document = PresentationDocument.Open(stream, true))
        {
            var properties = document.PresentationPart!.SlideParts.Single().Slide.Descendants<P.NonVisualDrawingProperties>().Last();
            var id = properties.GetFirstChild<A.HyperlinkOnClick>()!.Id!.Value;
            properties.AddChild(new A.HyperlinkOnHover { Id = id }, true);
        }

        using var pres = new Presentation(stream);

        // Act
        pres.Slides[0].Shapes.Last().Link.Remove();

        // Assert
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var result = PresentationDocument.Open(saved, false);
        var part = result.PresentationPart!.SlideParts.Single();
        part.Slide.Descendants<A.HyperlinkOnClick>().Should().BeEmpty();
        part.Slide.Descendants<A.HyperlinkOnHover>().Should().HaveCount(1);
        part.HyperlinkRelationships.Should().HaveCount(1);
        ValidatePresentation(pres);
    }

    private static Presentation LinkPresentation()
    {
        var pres = new Presentation(p => p.Slide());
        pres.Slides[0].Shapes.AddTextBox(0, 0, 100, 50, "Button");
        return pres;
    }

    private static int RelationshipCount(Presentation pres)
    {
        using var stream = new MemoryStream();
        pres.Save(stream);
        using var document = PresentationDocument.Open(stream, false);
        return document.PresentationPart!.SlideParts.Sum(part => part.HyperlinkRelationships.Count());
    }
}
