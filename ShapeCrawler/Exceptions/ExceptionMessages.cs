using ShapeCrawler.Charts;

namespace ShapeCrawler.Exceptions
{
    /// <summary>
    ///     Contains constant error messages.
    /// </summary>
    public static class ExceptionMessages
    {
        public const string NotTitle = "Chart has not a title.";

        /// <summary>
        ///     Returns message string with placeholder.
        /// </summary>
        public const string PresentationIsLarge = "The size of presentation more than {0} bytes.";

        public const string SlidesMuchMore = "The number of slides is more allowed {0}.";

        public const string PropertyCanChangedInNextVersion =
            "This property can not be changed for placeholder. The capability was planned to implement in one of the next library version. Use can use IsPlaceholder to check whether the shape is a placeholder.";

        public const string ForGroupedCanNotChanged =
            "This property can not be changed for a grouped shape. Use IsGrouped to check whether the shape is grouped.";

        public static string SeriesHasNotName =>
            $"The Series does not have a name. Use {nameof(Series.HasName)} to check whether series has a name.";

        public static string NotXValues =>
            $"This chart type has not {nameof(SlideChart.XValues)} property. You can check it via {nameof(SlideChart.HasXValues)} property.";

        public static string ChartCanNotHaveCategory =>
            $"#0 can not have category. You can check chart type via {nameof(SlideChart.Type)} property.";
    }
}