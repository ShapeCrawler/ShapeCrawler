using System;

namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft bubble chart shape.
/// </summary>
public sealed class DraftBubbleChartShape
{
    internal int ShapeX { get; private set; } = 100;

    internal int ShapeY { get; private set; } = 100;

    internal int ShapeWidth { get; private set; } = 400;

    internal int ShapeHeight { get; private set; } = 300;

    internal DraftBubbleChart? DraftBubbleChartBuilder { get; private set; }

    /// <summary>
    ///     Sets X-position in points.
    /// </summary>
    public DraftBubbleChartShape X(int x)
    {
        this.ShapeX = x;
        return this;
    }

    /// <summary>
    ///     Sets Y-position in points.
    /// </summary>
    public DraftBubbleChartShape Y(int y)
    {
        this.ShapeY = y;
        return this;
    }

    /// <summary>
    ///     Sets width in points.
    /// </summary>
    public DraftBubbleChartShape Width(int width)
    {
        this.ShapeWidth = width;
        return this;
    }

    /// <summary>
    ///     Sets height in points.
    /// </summary>
    public DraftBubbleChartShape Height(int height)
    {
        this.ShapeHeight = height;
        return this;
    }

    /// <summary>
    ///     Configures the bubble chart using a nested builder.
    /// </summary>
    /// <param name="configure">Delegate that configures the draft bubble chart builder instance.</param>
    public DraftBubbleChartShape Chart(Action<DraftBubbleChart> configure)
    {
        var builder = new DraftBubbleChart();
        configure(builder);
        this.DraftBubbleChartBuilder = builder;
        return this;
    }
}