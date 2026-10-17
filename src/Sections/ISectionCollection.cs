using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a collection of presentation sections.
/// </summary>
public interface ISectionCollection : IReadOnlyCollection<ISection>
{
    /// <summary>
    ///     Gets the section by index.
    /// </summary>
    ISection this[int index] { get; }

    /// <summary>
    ///     Adds a section starting at the specified one-based slide number and ending before the next section.
    ///     When adding the first section after slide 1, preceding slides are placed in a Default Section.
    /// </summary>
    /// <param name="name">The section name.</param>
    /// <param name="firstSlideNumber">The one-based number of the first slide in the section.</param>
    /// <returns>The added section.</returns>
    ISection Add(string name, int firstSlideNumber);

    /// <summary>
    ///     Removes specified section.
    /// </summary>
    void Remove(ISection removingSection);

    /// <summary>
    ///     Gets section by section name.
    /// </summary>
    ISection GetByName(string sectionName);
}

internal sealed class SectionCollection(PresentationDocument presDocument) : ISectionCollection
{
    public int Count => this.SectionList().Count;

    public ISection this[int index] => this.SectionList()[index];

    public IEnumerator<ISection> GetEnumerator() => this.SectionList().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    public ISection Add(string name, int firstSlideNumber)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        var presentation = presDocument.PresentationPart!.Presentation!;
        var slideIds = presentation.SlideIdList?.Elements<P.SlideId>().ToList() ?? [];
        if (firstSlideNumber < 1 || firstSlideNumber > slideIds.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(firstSlideNumber));
        }

        var sectionList = this.EnsureSectionList();
        if (!sectionList.Elements<P14.Section>().Any() && firstSlideNumber > 1)
        {
            sectionList.Append(CreateSection("Default Section", slideIds.Take(firstSlideNumber - 1)));
        }

        var positions = slideIds.Select((slide, index) => (Id: slide.Id!.Value, Index: index))
            .ToDictionary(item => item.Id, item => item.Index);
        var nextSection = sectionList.Elements<P14.Section>().FirstOrDefault(section =>
            section.SectionSlideIdList?.Elements<P14.SectionSlideIdListEntry>().FirstOrDefault() is { } first &&
            positions[first.Id!.Value] > firstSlideNumber - 1);
        var endIndex = nextSection == null
            ? slideIds.Count
            : positions[nextSection.SectionSlideIdList!.Elements<P14.SectionSlideIdListEntry>().First().Id!.Value];
        var sectionSlides = slideIds.Skip(firstSlideNumber - 1).Take(endIndex - firstSlideNumber + 1).ToList();
        var movedIds = sectionSlides.Select(slide => slide.Id!.Value).ToHashSet();
        foreach (var entry in sectionList.Descendants<P14.SectionSlideIdListEntry>()
                     .Where(entry => movedIds.Contains(entry.Id!.Value)).ToList())
        {
            entry.Remove();
        }

        var addedSection = CreateSection(name, sectionSlides);
        sectionList.InsertBefore(addedSection, nextSection);
        return new Section(addedSection);
    }

    public void Remove(ISection removingSection)
    {
        if (removingSection is not IRemovable removeable)
        {
            throw new SCException("Section cannot be removed.");
        }

        var total = this.Count;
        removeable.Remove();

        if (total == 1)
        {
            presDocument.PresentationPart!.Presentation!.PresentationExtensionList
                ?.Descendants<P14.SectionList>().First()
                .Remove();
        }
    }

    public ISection GetByName(string sectionName) => this.SectionList().First(section => section.Name == sectionName);

    private static P14.Section CreateSection(string name, IEnumerable<P.SlideId> slides)
    {
        return new P14.Section(new P14.SectionSlideIdList(slides.Select(slide =>
            new P14.SectionSlideIdListEntry { Id = slide.Id!.Value })))
        {
            Name = name,
            Id = Guid.NewGuid().ToString("B").ToUpperInvariant()
        };
    }

    private P14.SectionList EnsureSectionList()
    {
        var presentation = presDocument.PresentationPart!.Presentation!;
        presentation.PresentationExtensionList ??= new P.PresentationExtensionList();
        var extensionList = presentation.PresentationExtensionList;
        var sectionList = extensionList.Descendants<P14.SectionList>().FirstOrDefault();
        if (sectionList != null)
        {
            return sectionList;
        }

        const string uri = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
        var extension = extensionList.Elements<P.PresentationExtension>()
            .FirstOrDefault(item => item.Uri?.Value == uri);
        extension ??= extensionList.AppendChild(new P.PresentationExtension { Uri = uri });

        return extension.AppendChild(new P14.SectionList());
    }

    private List<Section> SectionList()
    {
        var p14SectionList = presDocument.PresentationPart!.Presentation!.PresentationExtensionList
            ?.Descendants<P14.SectionList>().FirstOrDefault();
        return p14SectionList == null
            ? []
            : [.. p14SectionList.OfType<P14.Section>().Select(p14Section => new Section(p14Section))];
    }
}
