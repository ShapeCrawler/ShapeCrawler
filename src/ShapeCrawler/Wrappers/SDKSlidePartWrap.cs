using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Wrappers;

internal sealed record SDKSlidePartWrap
{
    private readonly SlidePart sdkSlidePart;

    internal SDKSlidePartWrap(SlidePart sdkSlidePart)
    {
        this.sdkSlidePart = sdkSlidePart;
    }

    internal PresentationDocument SDKPresentationDocument()
    {
        return (PresentationDocument)this.sdkSlidePart.OpenXmlPackage;
    }
}