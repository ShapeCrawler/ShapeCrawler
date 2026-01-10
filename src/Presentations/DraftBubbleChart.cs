namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft bubble chart.
/// </summary>
public sealed class DraftBubbleChart
{
    internal string ChartName { get; private set; } = "Bubble Chart";

    internal string SeriesName { get; private set; } = "Series 1";

    internal (double X, double Y, double Size)[] SeriesPoints { get; private set; } = [];

    /// <summary>
    ///     Sets chart name.
    /// </summary>
    /// <param name="name">The name (title or identifier) to assign to the bubble chart.</param>
    public DraftBubbleChart Name(string name)
    {
        this.ChartName = name;
        return this;
    }

    /// <summary>
    ///     Adds a series with points to the bubble chart.
    /// </summary>
    /// <param name="seriesName">The name of the data series to be added to the bubble chart.</param>
    /// <param name="points">The points for the series.</param>
    public DraftBubbleChart Series(string seriesName, params (double X, double Y, double Size)[] points)
    {
        this.SeriesName = seriesName;
        this.SeriesPoints = points ?? throw new SCException($"{nameof(points)} cannot be null.");
        return this;
    }
}