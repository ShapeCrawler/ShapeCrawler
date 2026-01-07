using System;

namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft clustered bar chart shape.
/// </summary>
public sealed class DraftClusteredBarChartShape
{
    internal int ShapeX { get; private set; } = 100;

    internal int ShapeY { get; private set; } = 100;

    internal int ShapeWidth { get; private set; } = 400;

    internal int ShapeHeight { get; private set; } = 300;

    internal DraftChart? DraftChartBuilder { get; private set; }

    /// <summary>
    ///     Sets X-position in points.
    /// </summary>
    public DraftClusteredBarChartShape X(int x)
    {
        this.ShapeX = x;
        return this;
    }

    /// <summary>
    ///     Sets Y-position in points.
    /// </summary>
    public DraftClusteredBarChartShape Y(int y)
    {
        this.ShapeY = y;
        return this;
    }

    /// <summary>
    ///     Sets width in points.
    /// </summary>
    public DraftClusteredBarChartShape Width(int width)
    {
        this.ShapeWidth = width;
        return this;
    }

    /// <summary>
    ///     Sets height in points.
    /// </summary>
    public DraftClusteredBarChartShape Height(int height)
    {
        this.ShapeHeight = height;
        return this;
    }

    /// <summary>
    ///     Configures the chart using a nested builder.
    /// </summary>
    /// <param name="configure">Delegate that configures the draft chart builder instance.</param>
    public DraftClusteredBarChartShape Chart(Action<DraftChart> configure)
    {
        var builder = new DraftChart();
        configure(builder);
        this.DraftChartBuilder = builder;
        return this;
    }
}