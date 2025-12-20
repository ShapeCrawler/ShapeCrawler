namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft indentation for fluent API.
/// </summary>
public sealed class DraftIndentation
{
    /// <summary>
    ///    Gets indentation before text in points.
    /// </summary>
    internal decimal? BeforeTextPoints { get; private set; }

    /// <summary>
    ///     Sets indentation before text in points.
    /// </summary>
    public DraftIndentation BeforeText(decimal points)
    {
        this.BeforeTextPoints = points;
        return this;
    }
}
