namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft pie chart builder.
/// </summary>
public sealed class DraftPieChart
{
    internal string ChartName { get; private set; } = "Pie Chart";

    internal int ChartX { get; private set; } = 100;

    internal int ChartY { get; private set; } = 100;

    internal int ChartWidth { get; private set; } = 400;

    internal int ChartHeight { get; private set; } = 300;

    internal string[] CategoryNames { get; private set; } = [];

    internal string SeriesName { get; private set; } = "Series 1";

    internal double[] SeriesValues { get; private set; } = [];

    /// <summary>
    ///     Sets chart name.
    /// </summary>
    public DraftPieChart Name(string name)
    {
        this.ChartName = name;
        return this;
    }

    /// <summary>
    ///     Sets X-position in points.
    /// </summary>
    public DraftPieChart X(int x)
    {
        this.ChartX = x;
        return this;
    }

    /// <summary>
    ///     Sets Y-position in points.
    /// </summary>
    public DraftPieChart Y(int y)
    {
        this.ChartY = y;
        return this;
    }

    /// <summary>
    ///     Sets width in points.
    /// </summary>
    public DraftPieChart Width(int width)
    {
        this.ChartWidth = width;
        return this;
    }

    /// <summary>
    ///     Sets height in points.
    /// </summary>
    public DraftPieChart Height(int height)
    {
        this.ChartHeight = height;
        return this;
    }

    /// <summary>
    ///     Sets categories for the pie chart.
    /// </summary>
    public DraftPieChart Categories(params string[] categories)
    {
        this.CategoryNames = categories ?? throw new SCException($"{nameof(categories)} cannot be null.");
        return this;
    }

    /// <summary>
    ///     Adds a series with values to the pie chart.
    /// </summary>
    public DraftPieChart Series(string seriesName, params double[] values)
    {
        this.SeriesName = seriesName;
        this.SeriesValues = values;
        return this;
    }
}