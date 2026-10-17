using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using ShapeCrawler.Tests.Helpers;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler.Tests;

[TestFixture]
public class SectionTests : SCTest
{
    [Test]
    public void Add_creates_section_with_slides_and_valid_xml()
    {
        // Arrange
        using var presentation = CreatePresentation();

        // Act
        var section = presentation.Sections.Add("Introduction & overview", 1);

        // Assert
        section.Name.Should().Be("Introduction & overview");
        section.Slides.Select(slide => slide.Number).Should().Equal(1, 2, 3, 4);
        using var reopened = SaveAndOpenPresentation(presentation);
        reopened.Sections.Should().ContainSingle();
        reopened.Sections[0].Name.Should().Be(section.Name);
        reopened.Sections[0].Slides.Select(slide => slide.Number).Should().Equal(1, 2, 3, 4);
        using var sdk = SaveAndOpenPresentationAsSdk(presentation);
        var sectionXml = sdk.PresentationPart!.Presentation.Descendants<P14.Section>().Single();
        Guid.TryParse(sectionXml.Id!.Value, out _).Should().BeTrue();
        sectionXml.Parent!.Parent!.GetAttribute("uri", "").Value
            .Should().Be("{521415D9-36F7-43E2-AB2F-B90AF26B5E84}");
        new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(sdk).Should().BeEmpty();
    }

    [Test]
    public void Add_creates_default_section_for_preceding_slides()
    {
        // Arrange
        using var presentation = CreatePresentation();

        // Act
        presentation.Sections.Add("Conclusion", 4);

        // Assert
        using var reopened = SaveAndOpenPresentation(presentation);
        reopened.Sections.Select(section => section.Name).Should().Equal("Default Section", "Conclusion");
        reopened.Sections[0].Slides.Select(slide => slide.Number).Should().Equal(1, 2, 3);
        reopened.Sections[1].Slides.Select(slide => slide.Number).Should().Equal(4);
        ValidatePresentation(reopened);
    }

    [Test]
    public void Add_splits_existing_section_and_preserves_following_section()
    {
        // Arrange
        using var presentation = CreatePresentation();
        var introduction = presentation.Sections.Add("Introduction", 1);
        presentation.Sections.Add("Conclusion", 4);

        // Act
        var content = presentation.Sections.Add("Content", 2);

        // Assert
        introduction.Slides.Select(slide => slide.Number).Should().Equal(1);
        content.Slides.Select(slide => slide.Number).Should().Equal(2, 3);
        content.Slides[0].Number.Should().Be(2);
        using var reopened = SaveAndOpenPresentation(presentation);
        reopened.Sections.Select(section => section.Name).Should().Equal("Introduction", "Content", "Conclusion");
        reopened.Sections[0].Slides.Select(slide => slide.Number).Should().Equal(1);
        reopened.Sections[1].Slides.Select(slide => slide.Number).Should().Equal(2, 3);
        reopened.Sections[2].Slides.Select(slide => slide.Number).Should().Equal(4);
        using var sdk = SaveAndOpenPresentationAsSdk(reopened);
        var sections = sdk.PresentationPart!.Presentation.Descendants<P14.Section>().ToList();
        sections.Select(section => section.Id!.Value).Should().OnlyHaveUniqueItems();
        sections.SelectMany(section => section.Descendants<P14.SectionSlideIdListEntry>())
            .Select(entry => entry.Id!.Value).Should().OnlyHaveUniqueItems().And.HaveCount(4);
        ValidatePresentation(reopened);
    }

    [Test]
    public void Add_uses_slide_order_after_slides_are_reordered()
    {
        // Arrange
        using var presentation = CreatePresentation();
        presentation.Slide(4).Number = 2;

        // Act
        presentation.Sections.Add("Content", 2);

        // Assert
        using var reopened = SaveAndOpenPresentation(presentation);
        reopened.Sections[0].Slides.Select(slide => slide.Number).Should().Equal(1);
        reopened.Sections[1].Slides.Select(slide => slide.Number).Should().Equal(2, 3, 4);
        using var sdk = SaveAndOpenPresentationAsSdk(reopened);
        var slideIds = sdk.PresentationPart!.Presentation.SlideIdList!
            .Elements<DocumentFormat.OpenXml.Presentation.SlideId>().Skip(1).Select(slide => slide.Id!.Value);
        sdk.PresentationPart.Presentation.Descendants<P14.Section>().Last()
            .Descendants<P14.SectionSlideIdListEntry>().Select(entry => entry.Id!.Value).Should().Equal(slideIds);
    }

    [Test]
    public void Add_at_existing_boundary_leaves_previous_section_empty()
    {
        // Arrange
        using var presentation = CreatePresentation();
        var original = presentation.Sections.Add("Original", 1);

        // Act
        presentation.Sections.Add("Replacement", 1);

        // Assert
        original.Slides.Should().BeEmpty();
        using var reopened = SaveAndOpenPresentation(presentation);
        reopened.Sections.Select(section => section.Name).Should().Equal("Original", "Replacement");
        reopened.Sections[0].Slides.Should().BeEmpty();
        reopened.Sections[1].Slides.Should().HaveCount(4);
        ValidatePresentation(reopened);
    }

    [Test]
    public void Add_can_recreate_sections_after_removing_last_section()
    {
        // Arrange
        using var presentation = CreatePresentation();
        var original = presentation.Sections.Add("Original", 1);
        presentation.Sections.Remove(original);

        // Act
        presentation.Sections.Add("New", 1);

        // Assert
        using var reopened = SaveAndOpenPresentation(presentation);
        reopened.Sections.Should().ContainSingle().Which.Name.Should().Be("New");
        using var sdk = SaveAndOpenPresentationAsSdk(reopened);
        sdk.PresentationPart!.Presentation.PresentationExtensionList!
            .Elements<DocumentFormat.OpenXml.Presentation.PresentationExtension>()
            .Count(extension => extension.Uri?.Value == "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}")
            .Should().Be(1);
        ValidatePresentation(reopened);
    }

    [TestCase(0)]
    [TestCase(-1)]
    [TestCase(5)]
    public void Add_rejects_invalid_slide_number_without_adding_sections(int firstSlideNumber)
    {
        // Arrange
        using var presentation = CreatePresentation();

        // Act
        var act = () => presentation.Sections.Add("Content", firstSlideNumber);

        // Assert
        act.Should().Throw<ArgumentOutOfRangeException>();
        presentation.Sections.Should().BeEmpty();
    }

    [TestCase(null)]
    [TestCase("")]
    [TestCase("  ")]
    public void Add_rejects_missing_name_without_adding_sections(string name)
    {
        // Arrange
        using var presentation = CreatePresentation();

        // Act
        var act = () => presentation.Sections.Add(name, 1);

        // Assert
        act.Should().Throw<ArgumentException>();
        presentation.Sections.Should().BeEmpty();
    }

    [Test]
    public void Added_section_tracks_slide_removal()
    {
        // Arrange
        using var presentation = CreatePresentation();
        var section = presentation.Sections.Add("Content", 3);

        // Act
        section.Slides[0].Remove();

        // Assert
        section.Slides.Select(slide => slide.Number).Should().Equal(3);
        using var reopened = SaveAndOpenPresentation(presentation);
        reopened.Sections[1].Slides.Select(slide => slide.Number).Should().Equal(3);
        ValidatePresentation(reopened);
    }

    private static Presentation CreatePresentation()
    {
        return new Presentation(p =>
        {
            p.Slide();
            p.Slide();
            p.Slide();
            p.Slide();
        });
    }
}
