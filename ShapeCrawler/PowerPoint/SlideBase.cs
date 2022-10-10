using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Services;

namespace ShapeCrawler;

internal abstract class SlideBase : IRemovable, IPresentationComponent
{
    public abstract bool IsRemoved { get; set; } // TODO: make internal

    public abstract SCPresentation PresentationInternal { get; }
    
    internal abstract TypedOpenXmlPart TypedOpenXmlPart { get; }

    public abstract void ThrowIfRemoved(); // TODO: make internal
}