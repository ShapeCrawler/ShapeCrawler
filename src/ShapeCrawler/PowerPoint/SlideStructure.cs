using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler;

internal abstract class SlideStructure : ISlideStructure
{
    protected SlideStructure(IPresentation pres)
    {
        this.Presentation = pres;
    }

    public IPresentation Presentation { get; protected init; }
    
    public abstract int Number { get; set; }

    internal SCPresentation PresentationInternal => (SCPresentation)this.Presentation;

    internal abstract TypedOpenXmlPart TypedOpenXmlPart { get; }
}