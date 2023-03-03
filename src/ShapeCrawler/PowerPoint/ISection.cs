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
    ISectionSlideCollection Slides { get; }

    /// <summary>
    ///     Gets section name.
    /// </summary>
    string Name { get; }
}

internal sealed class SCSection : ISection
{
    internal SCSection(SCSectionCollection sections, P14.Section p14Section)
    {
        this.Sections = sections;
        this.SDKSection = p14Section;
    }

    public ISectionSlideCollection Slides => new SCSectionSlideCollection(this);

    public string Name => this.GetName();

    internal SCSectionCollection Sections { get; }

    internal P14.Section SDKSection { get; }

    private string GetName()
    {
        return this.SDKSection.Name!;
    }
}