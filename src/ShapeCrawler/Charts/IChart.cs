using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using C = DocumentFormat.OpenXml.Drawing.Charts;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

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
    ///     Gets a value indicating whether the chart has a title.
    /// </summary>
    public bool HasTitle { get; }
    
    /// <summary>
    ///     Gets underlying instance of <see cref="DocumentFormat.OpenXml.Presentation.GraphicFrame"/>.
    /// </summary>
    public P.GraphicFrame SDKGraphicFrame { get; }
    
    /// <summary>
    ///     Gets underlying instance of <see cref="DocumentFormat.OpenXml.Packaging.ChartPart"/>.
    /// </summary>
    public ChartPart SDKChartPart { get; }
    
    /// <summary>
    ///     Gets underlying instance of <see cref="DocumentFormat.OpenXml.Drawing.Charts.PlotArea"/>.
    /// </summary>
    public C.PlotArea SDKPlotArea { get; }
    
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