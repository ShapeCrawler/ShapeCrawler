
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Wrappers;

internal sealed record SDKPPresentationWrap
{
    private readonly P.Presentation pPresentation;

    internal SDKPPresentationWrap(P.Presentation pPresentation)
    {
        this.pPresentation = pPresentation;
    }
}