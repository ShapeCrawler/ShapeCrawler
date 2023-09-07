using ShapeCrawler.Charts;

namespace ShapeCrawler.Exceptions;

internal static class ExceptionMessages
{
    internal const string PresentationIsLarge = "The size of presentation more than {0} bytes.";

    internal const string SlidesMuchMore = "The number of slides is more allowed {0}.";

    internal static string SeriesHasNotName =>
        $"The Series does not have a name. Use {nameof(Series.HasName)} to check whether series has a name.";

    internal static string NotXValues =>
        $"This chart type has not {nameof(SlideChart.XValues)} property. You can check it via {nameof(SlideChart.HasXValues)} property.";
}