using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Services;

namespace ShapeCrawler;

/// <summary>
///     Represents slide object.
/// </summary>
public interface ISlideObject : IPresentationComponent
{

}

internal abstract class SlideObject : ISlideObject
{
    protected SlideObject(IPresentation pres)
    {
        this.Presentation = pres;
    }

    public IPresentation Presentation { get; }

    internal SCPresentation PresentationInternal => (SCPresentation)this.Presentation;

    internal abstract TypedOpenXmlPart TypedOpenXmlPart { get; }
}