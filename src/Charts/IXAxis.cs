namespace ShapeCrawler.Charts;

/// <summary>
///     Represents a chart X-axis.
/// </summary>
public interface IXAxis
{
    /// <summary>
    ///     Gets axis values.
    /// </summary>
    double[] Values { get; }
    
    /// <summary>
    ///     Gets or sets axis minimum value.
    /// </summary>
    double Minimum { get; set; }
    
    /// <summary>
    ///     Gets or sets axis maximum value.
    /// </summary>
    double Maximum { get; set; }
}