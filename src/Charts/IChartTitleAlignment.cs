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
}