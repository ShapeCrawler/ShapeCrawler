#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a chart point.
/// </summary>
public interface IChartPoint
{
    /// <summary>
    ///     Gets or sets chart point value.
    /// </summary>
    public double Value { get; set; }
}