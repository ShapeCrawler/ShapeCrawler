using ShapeCrawler.Collections;
using ShapeCrawler.Shapes;

// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a chart.
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
        ///     Gets a value indicating whether the chart has a title.
        /// </summary>
        public bool HasTitle { get; }

        /// <summary>
        ///     Gets a value indicating whether the chart has categories.
        /// </summary>
        /// <remarks>Some chart types like ScatterChart and BubbleChart does not have categories.</remarks>
        bool HasCategories { get; }

        /// <summary>
        ///     Gets collection of the chart series.
        /// </summary>
        ISeriesCollection SeriesCollection { get; }

        /// <summary>
        ///     Gets collection of chart categories.
        /// </summary>
        CategoryCollection Categories { get; }

        /// <summary>
        ///     Gets a value indicating whether the chart has x-axis values.
        /// </summary>
        bool HasXValues { get; }

        /// <summary>
        ///     Gets collection of x-axis values.
        /// </summary>
        LibraryCollection<double> XValues { get; }

        byte[] SpreadsheetByteArray { get; }
    }
}