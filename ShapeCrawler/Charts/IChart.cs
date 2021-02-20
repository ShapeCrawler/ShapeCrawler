using ShapeCrawler.Collections;
using ShapeCrawler.Models;

namespace ShapeCrawler.Charts
{
    /// <summary>
    ///     Represents a chart on  a slide.
    /// </summary>
    public interface IChart : IShape
    {
        /// <summary>
        ///     Gets chart type.
        /// </summary>
        ChartType Type { get; }

        /// <summary>
        ///     Gets the chart title. Returns null if the chart has not a title.
        /// </summary>
        string Title { get; }

        /// <summary>
        ///     Determines whether the chart has a title.
        /// </summary>
        public bool HasTitle { get; }

        /// <summary>
        ///     Determines whether the chart has categories.
        /// </summary>
        /// <remarks>Some chart types like ScatterChart and BubbleChart does not have categories.</remarks>
        bool HasCategories { get; }

        /// <summary>
        ///     Gets collection of the chart series.
        /// </summary>
        SeriesCollection SeriesCollection { get; }

        /// <summary>
        ///     Gets collection of the chart category.
        /// </summary>
        CategoryCollection Categories { get; }

        /// <summary>
        ///     Determines whether the chart has x-axis values.
        /// </summary>
        bool HasXValues { get; }

        /// <summary>
        ///     Gets collection of x-axis values.
        /// </summary>
        LibraryCollection<double> XValues { get; }
    }
}