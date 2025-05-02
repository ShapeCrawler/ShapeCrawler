using System.Collections.Generic;

// ReSharper disable once CheckNamespace
#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a chart.
/// </summary>
public interface IChart : IShape
{
    /// <summary>
    ///     Gets the chart type.
    /// </summary>
    ChartType Type { get; }

    /// <summary>
    ///     Gets chart title. Returns <c>null</c> if the chart does not have a title.
    /// </summary>
    string? Title { get; }

    /// <summary>
    ///     Gets a value indicating whether the chart has categories.
    /// </summary>
    bool HasCategories { get; }

    /// <summary>
    ///     Gets the collection of categories.
    /// </summary>
    public IReadOnlyList<ICategory> Categories { get; }

    /// <summary>
    ///     Gets collection of data series.
    /// </summary>
    ISeriesCollection SeriesCollection { get; }

    /// <summary>
    ///     Gets the byte array of the underlying data source spreadsheet.
    /// </summary>
    byte[] GetSpreadsheetByteArray();

    /// <summary>
    ///     Represents the chart as a Scatter.  
    /// </summary>
    /// <exception cref="SCException">Thrown if the chart is not a Scatter.</exception>
    IScatterChart AsScatterChart();
}