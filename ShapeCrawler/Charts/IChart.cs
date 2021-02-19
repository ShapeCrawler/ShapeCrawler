using ShapeCrawler.Collections;
using ShapeCrawler.Models;

namespace ShapeCrawler.Charts
{
    public interface IChart : IShape
    {
        /// <summary>
        ///     Gets the chart title. Returns null if chart has not a title.
        /// </summary>
        ChartType Type { get; }

        /// <summary>
        ///     Gets chart title string.
        /// </summary>
        string Title { get; }

        /// <summary>
        ///     Determines whether chart has a title.
        /// </summary>
        public bool HasTitle { get; }


        /// <summary>
        ///     Determines whether chart has categories. Some chart types like ScatterChart and BubbleChart does not have
        ///     categories.
        /// </summary>
        bool HasCategories { get; }

        /// <summary>
        ///     Gets collection of the chart series.
        /// </summary>
        SeriesCollection SeriesCollection { get; }

        /// <summary>
        ///     Gets collection of the chart category.
        /// </summary>
        CategoryCollection Categories { get; }

        bool HasXValues { get; }

        LibraryCollection<double> XValues { get; }
    }
}