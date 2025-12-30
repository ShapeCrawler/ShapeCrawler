using System;

namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft pie chart shape builder.
/// </summary>
public sealed class DraftPieChartShape
{
    internal int ShapeX { get; private set; } = 100;

    internal int ShapeY { get; private set; } = 100;

    internal int ShapeWidth { get; private set; } = 400;

    internal int ShapeHeight { get; private set; } = 300;

    internal DraftPieChart? DraftPieChartBuilder { get; private set; }

    /// <summary>
    ///     Sets X-position in points.
    /// </summary>
    public DraftPieChartShape X(int x)
    {
        this.ShapeX = x;
        return this;
    }

    /// <summary>
    ///     Sets Y-position in points.
    /// </summary>
    public DraftPieChartShape Y(int y)
    {
        this.ShapeY = y;
        return this;
    }

    /// <summary>
    ///     Sets width in points.
    /// </summary>
    public DraftPieChartShape Width(int width)
    {
        this.ShapeWidth = width;
        return this;
    }

    /// <summary>
    ///     Sets height in points.
    /// </summary>
    public DraftPieChartShape Height(int height)
    {
        this.ShapeHeight = height;
        return this;
    }

    /// <summary>
    ///     Configures the pie chart using a nested builder.
    /// </summary>
    /// <param name="configure">Delegate that configures the draft pie chart builder instance.</param>
    public DraftPieChartShape Chart(Action<DraftPieChart> configure)
    {
        var builder = new DraftPieChart();
        configure(builder);
        this.DraftPieChartBuilder = builder;
        return this;
    }
}