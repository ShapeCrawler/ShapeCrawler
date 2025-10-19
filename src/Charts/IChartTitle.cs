#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a chart title.
/// </summary>
public interface IChartTitle
{
    /// <summary>
    ///     Gets or sets the title text.
    /// </summary>
    string? Text { get; set; }

    /// <summary>
    ///     Gets or sets the font color as hexadecimal representation.
    /// </summary>
    string FontColor { get; set; }

    /// <summary>
    ///     Gets or sets the font size in points.
    /// </summary>
    int FontSize { get; set; }
}
