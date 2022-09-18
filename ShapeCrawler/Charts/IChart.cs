using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;
using ShapeCrawler.Shapes;

// ReSharper disable once CheckNamespace
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
        ///     Gets title. Returns null if chart doesn't have title.
        /// </summary>
        string Title { get; }

        /// <summary>
        ///     Gets a value indicating whether the chart has a title.
        /// </summary>
        public bool HasTitle { get; }

        /// <summary>
        ///     Gets a value indicating whether the chart type has categories.
        /// </summary>
        bool HasCategories { get; }

        /// <summary>
        ///     Gets category collection. Returns null if the chart type doesn't have categories.
        /// </summary>
        public ICategoryCollection Categories { get; }

        /// <summary>
        ///     Gets collection of data series.
        /// </summary>
        ISeriesCollection SeriesCollection { get; }

        /// <summary>
        ///     Gets a value indicating whether the chart has x-axis values.
        /// </summary>
        bool HasXValues { get; }

        /// <summary>
        ///     Gets collection of x-axis values.
        /// </summary>
        LibraryCollection<double> XValues { get; } // TODO: should be excluded

        /// <summary>
        ///     Gets byte array of workbook containing chart data source.
        /// </summary>
        byte[] WorkbookByteArray { get; }

        /// <summary>
        ///     Gets parent slide.
        /// </summary>
        ISlide ParentSlide { get; }

        /// <summary>
        ///     Gets instance of <see cref="SpreadsheetDocument"/> of Open XML SDK.
        /// </summary>
        SpreadsheetDocument SDKSpreadsheetDocument { get; }
    }
}