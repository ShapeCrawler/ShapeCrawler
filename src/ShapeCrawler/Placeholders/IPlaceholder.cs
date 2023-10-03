// ReSharper disable CheckNamespace

using ShapeCrawler.Placeholders;

namespace ShapeCrawler;

/// <summary>
///     Represents a placeholder.
/// </summary>
public interface IPlaceholder
{
    /// <summary>
    ///     Gets placeholder type.
    /// </summary>
    PlaceholderType Type { get; }
}