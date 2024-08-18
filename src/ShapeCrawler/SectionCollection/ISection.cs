using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.SectionCollection;
using ShapeCrawler.ShapeCollection;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

// ReSharper disable once CheckNamespace
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
            new SectionSlides(sdkPresDocument, p14Section.Descendants<P14.SectionSlideIdListEntry>()))
    {
    }

    private Section(P14.Section p14Section, IReadOnlyList<ISlide> slides)
    {
        this.P14Section = p14Section;
        this.Slides = slides;
    }

    public string Name => this.GetName();
    
    public IReadOnlyList<ISlide> Slides { get; }
    
    private P14.Section P14Section { get; }
    
    public void Remove() => this.P14Section.Remove();

    private string GetName() => this.P14Section.Name!;
}