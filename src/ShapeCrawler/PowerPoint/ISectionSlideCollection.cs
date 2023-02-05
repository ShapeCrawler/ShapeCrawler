using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler;

/// <summary>
///     Represents collection of section slide.
/// </summary>
public interface ISectionSlideCollection : IEnumerable<ISlide>
{
    /// <summary>
    ///     Gets sections count.
    /// </summary>
    int Count { get; }

    /// <summary>
    ///     Gets section slide by index.
    /// </summary>
    ISlide this[int index] { get; }
}

internal sealed class SCSectionCollection : ISectionCollection
{
    private readonly List<SCSection> sections;
    private readonly SectionList? sdkSectionList;

    private SCSectionCollection(SCPresentation presentation, List<SCSection> sections)
    {
        this.Presentation = presentation;
        this.sections = sections;
    }

    private SCSectionCollection(SCPresentation presentation, List<SCSection> sections, SectionList sdkSectionList)
    {
        this.Presentation = presentation;
        this.sections = sections;
        this.sdkSectionList = sdkSectionList;
    }

    public int Count => this.sections.Count;

    internal SCPresentation Presentation { get; }

    public ISection this[int i] => this.sections[i];

    public IEnumerator<ISection> GetEnumerator()
    {
        return this.sections.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.sections.GetEnumerator();
    }

    public void Remove(ISection removingSection)
    {
        if (this.sdkSectionList == null || this.Count == 0)
        {
            return;
        }

        var sectionInternal = (SCSection)removingSection;
        sectionInternal.SDKSection.Remove();

        if (this.sections.Count == 1)
        {
            this.sdkSectionList.Remove();
        }

        this.sections.Remove(sectionInternal);
    }

    public ISection GetByName(string sectionName)
    {
        return this.sections.First(section => section.Name == sectionName);
    }

    internal static SCSectionCollection Create(SCPresentation presentation)
    {
        var sections = new List<SCSection>();
        var sdkSectionList = presentation.SDKPresentationInternal.PresentationPart!.Presentation.PresentationExtensionList
            ?.Descendants<SectionList>().FirstOrDefault();

        if (sdkSectionList == null)
        {
            return new SCSectionCollection(presentation, sections);
        }

        var sectionCollection = new SCSectionCollection(presentation, sections, sdkSectionList);

        foreach (P14.Section sdkSection in sdkSectionList)
        {
            sections.Add(new SCSection(sectionCollection, sdkSection));
        }

        return new SCSectionCollection(presentation, sections, sdkSectionList);
    }

    internal void RemoveSldId(string id)
    {
        var removing = this.sdkSectionList?.Descendants<P14.SectionSlideIdListEntry>().FirstOrDefault(s => s.Id == id);
        if (removing == null)
        {
            return;
        }

        removing.Remove();
        this.Presentation.SDKPresentationInternal.PresentationPart!.Presentation.Save();
    }
}