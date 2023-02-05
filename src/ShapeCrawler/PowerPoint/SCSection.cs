using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler;

[SuppressMessage("ReSharper", "InconsistentNaming", Justification = "Exception")]
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

    internal Section SDKSection { get; }

    private string GetName()
    {
        return this.SDKSection.Name!;
    }
}