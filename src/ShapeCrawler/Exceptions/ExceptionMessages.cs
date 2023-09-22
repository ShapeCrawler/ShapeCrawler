using ShapeCrawler.Charts;

namespace ShapeCrawler.Exceptions;

internal static class ExceptionMessages
{
    internal static string NotXValues =>
        $"This chart type has not {nameof(SlideChart.XValues)} property. You can check it via {nameof(SlideChart.HasXValues)} property.";
}