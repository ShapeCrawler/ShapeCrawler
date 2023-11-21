using ShapeCrawler.Charts;

namespace ShapeCrawler.Exceptions;

internal static class ExceptionMessages
{
    internal static string NotXValues =>
        $"This chart type has not {nameof(Chart.XValues)} property. You can check it via {nameof(Chart.HasXValues)} property.";
}