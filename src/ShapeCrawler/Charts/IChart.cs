using System.Collections.Generic;

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
    ChartType Type { get; }

    /// <summary>
    ///     Gets a value indicating whether the chart has a title.
    /// </summary>
    public bool HasTitle { get; }
    
    /// <summary>
    ///     Gets chart title.
    /// </summary>
    string Title { get; }

    /// <summary>
    ///     Gets a value indicating whether the chart has categories.
    /// </summary>
    bool HasCategories { get; }

    /// <summary>
    ///     Gets collection of categories.
    /// </summary>
    public IReadOnlyList<ICategory> Categories { get; }

    /// <summary>
    ///     Gets collection of data series.
    /// </summary>
    ISeriesList SeriesList { get; }

    /// <summary>
    ///     Gets a value indicating whether the chart has x-axis values.
    /// </summary>
    bool HasXValues { get; }

    /// <summary>
    ///     Gets collection of x-axis values.
    /// </summary>
    List<double> XValues { get; } // TODO: should be excluded
    
    /// <summary>
    ///     Gets chart axes manager.
    /// </summary>
    IAxesManager Axes { get; }

    /// <summary>
    ///     Gets byte array of excel book containing chart data source.
    /// </summary>
    byte[] BookByteArray();
}