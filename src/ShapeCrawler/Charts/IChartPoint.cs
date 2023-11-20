// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

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