#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents the alignment properties of a chart title.
/// </summary>
public interface IChartTitleAlignment
{
    /// <summary>
    ///     Gets or sets the custom rotation angle of the chart title in degrees.
    /// </summary>
    decimal CustomAngle { get; set; }

    /// <summary>
    ///     Gets or sets the horizontal position of the chart title as a factor of the chart width (0.0 to 1.0).
    ///     A value of 0.0 represents the left edge, and 1.0 represents the right edge.
    ///     Returns null if the title is using automatic positioning.
    /// </summary>
    decimal? X { get; set; }

    /// <summary>
    ///     Gets or sets the vertical position of the chart title as a factor of the chart height (0.0 to 1.0).
    ///     A value of 0.0 represents the top edge, and 1.0 represents the bottom edge.
    ///     Returns null if the title is using automatic positioning.
    /// </summary>
    decimal? Y { get; set; }
}