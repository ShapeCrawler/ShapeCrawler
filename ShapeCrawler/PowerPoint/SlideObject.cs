using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Services;

namespace ShapeCrawler;

/// <summary>
///     Represents slide object.
/// </summary>
public interface ISlideObject : IPresentationComponent
{
    /// <summary>
    ///     Gets or sets slide number.
    /// </summary>
    int Number { get; set; }
}

internal abstract class SlideObject : ISlideObject
{
    protected SlideObject(IPresentation pres)
    {
        this.Presentation = pres;
    }

    public IPresentation Presentation { get; init; }
    
    public abstract int Number { get; set; }

    internal SCPresentation PresentationInternal => (SCPresentation)this.Presentation;

    internal abstract TypedOpenXmlPart TypedOpenXmlPart { get; }
}