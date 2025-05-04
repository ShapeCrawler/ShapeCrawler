using System.Collections.Generic;
using ShapeCrawler.Charts;

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
    ///     Gets chart type.
    /// </summary>
    ChartType Type { get; }

    /// <summary>
    ///     Gets title. Returns <c>null</c> if the chart doesn't have a title.
    /// </summary>
    string? Title { get; }

    /// <summary>
    ///     Gets category collection. Returns <c>null</c> if the chart type doesn't have categories, e.g., Scatter.
    /// </summary>
    public IReadOnlyList<ICategory>? Categories { get; }
    
    /// <summary>
    ///     Gets X-Axis. Returns <c>null</c> if the chart type doesn't have an X-Axis, e.g., Pie.
    /// </summary>
    IXAxis? XAxis { get; }

    /// <summary>
    ///     Gets series collection.
    /// </summary>
    ISeriesCollection SeriesCollection { get; }

    /// <summary>
    ///     Gets byte array of the chart data store worksheet.
    /// </summary>
    byte[] GetWorksheetByteArray();
}