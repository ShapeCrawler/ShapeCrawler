﻿using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Models.SlideComponents.Chart;

namespace ShapeCrawler.Exceptions
{
    /// <summary>
    /// Contains constant error messages.
    /// </summary>
    public static class ExceptionMessages
    {
        public const string NoTextFrame = "Element has not a text frame.";

        public const string NoChart = "Element has not a chart.";

        public const string NoPicture = "Element has not a picture.";

        public const string NoTable = "Element has not a table.";

        public const string NoOleObject = "Element has not a OLE object.";

        public const string NotTitle = "Chart has not a title.";

        public static string SeriesHasNotName => $"The Series does not have a name. Use {nameof(Series.HasName)} to check whether series has a name.";

        /// <summary>
        /// Returns message string with placeholder.
        /// </summary>
        public const string PresentationIsLarge = "The size of presentation more than {0} bytes.";

        public const string SlidesMuchMore = "The number of slides is more allowed {0}.";

        public const string PropertyCanChangedInNextVersion =
            "This property can not be changed for placeholder. The capability was planned to implement in one of the next library version. Use can use IsPlaceholder to check whether the shape is a placeholder.";

        public const string ForGroupedCanNotChanged =
            "This property can not be changed for a grouped shape. Use IsGrouped to check whether the shape is grouped.";

        public static string NotXValues =>
            $"This chart type has not {nameof(ChartSc.XValues)} property. You can check it via {nameof(ChartSc.HasXValues)} property.";

        public static string ChartCanNotHaveCategory =>
            $"#0 can not have category. You can check chart type via {nameof(ChartSc.Type)} property.";

        public static string ShapeIsNotPlaceholder =>
            $"The shape is not placeholder. You can check it via {nameof(ShapeSc.IsPlaceholder)} property.";
    }
}
