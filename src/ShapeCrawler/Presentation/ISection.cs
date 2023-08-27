using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler;

/// <summary>
///     Represents a presentation section.
/// </summary>
public interface ISection
{
    /// <summary>
    ///     Gets section slides.
    /// </summary>
    IReadOnlyList<ISlide> Slides { get; }

    /// <summary>
    ///     Gets section name.
    /// </summary>
    string Name { get; }
}

internal sealed class Section : ISection, IRemoveable
{
    internal Section(PresentationDocument sdkPresDocument, P14.Section p14Section)
        : this(
            p14Section,
            new SectionSlides(sdkPresDocument, p14Section.Descendants<P14.SectionSlideIdListEntry>())
        )
    {
    }

    private Section(P14.Section p14Section, IReadOnlyList<ISlide> slides)
    {
        this.p14Section = p14Section;
        this.Slides = slides;
    }

    public IReadOnlyList<ISlide> Slides { get; }

    public string Name => this.GetName();

    private P14.Section p14Section { get; }

    private string GetName()
    {
        return this.p14Section.Name!;
    }

    public void Remove() => this.p14Section.Remove();
}