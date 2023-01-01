namespace ShapeCrawler.Placeholders;

/// <summary>
///     Represents a placeholder.
/// </summary>
public interface IPlaceholder
{
    /// <summary>
    ///     Gets placeholder type.
    /// </summary>
    SCPlaceholderType Type { get; }
}