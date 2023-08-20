using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Placeholders;

internal sealed class SlidePlaceholder : IPlaceholder
{
    private readonly P.PlaceholderShape pPlaceholderShape;

    internal SlidePlaceholder(P.PlaceholderShape pPlaceholderShape)
    {
        this.pPlaceholderShape = pPlaceholderShape;
    }

    public SCPlaceholderType Type { get; }
}