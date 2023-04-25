using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Collections;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a chart.
/// </summary>
public interface IChart : IShape
{
    /// <summary>
    ///     Gets chart type.
    /// </summary>
    SCChartType Type { get; }

    /// <summary>
    ///     Gets chart title if chart has it, otherwise <see langword="null"/>.
    /// </summary>
    string? Title { get; }

    /// <summary>
    ///     Gets a value indicating whether the chart has a title.
    /// </summary>
    public bool HasTitle { get; }

    /// <summary>
    ///     Gets a value indicating whether the chart type has categories.
    /// </summary>
    bool HasCategories { get; }

    /// <summary>
    ///     Gets collection of categories. Returns <see langword="null"/> if chart type has no categories.
    /// </summary>
    public ICategoryCollection? Categories { get; }

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
    SCLibraryCollection<double> XValues { get; } // TODO: should be excluded

    /// <summary>
    ///     Gets byte array of workbook containing chart data source.
    /// </summary>
    byte[] WorkbookByteArray { get; }

    /// <summary>
    ///     Gets instance of <see cref="SpreadsheetDocument"/> of Open XML SDK.
    /// </summary>
    SpreadsheetDocument SDKSpreadsheetDocument { get; }

    /// <summary>
    ///     Gets chart axes manager.
    /// </summary>
    IAxesManager Axes { get; }
}