namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft stroke.
/// </summary>
public sealed class DraftStroke
{
    internal decimal? DraftWidthPoints { get; private set; }

    /// <summary>
    ///     Sets stroke width in points.
    /// </summary>
    public DraftStroke Width(decimal points)
    {
        this.DraftWidthPoints = points;
        return this;
    }
}