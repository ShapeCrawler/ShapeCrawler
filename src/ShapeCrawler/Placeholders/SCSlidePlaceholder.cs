using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Placeholders;

internal sealed class SCSlidePlaceholder : IPlaceholder
{
    private readonly P.PlaceholderShape pPlaceholderShape;

    internal SCSlidePlaceholder(P.PlaceholderShape pPlaceholderShape)
    {
        this.pPlaceholderShape = pPlaceholderShape;
    }

    public SCPlaceholderType Type { get; }
}