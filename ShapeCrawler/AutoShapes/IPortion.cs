// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a portion of a paragraph.
/// </summary>
public interface IPortion
{
    /// <summary>
    ///     Gets or sets text.
    /// </summary>
    string Text { get; set; }

    /// <summary>
    ///     Gets font.
    /// </summary>
    IFont Font { get; }

    /// <summary>
    ///     Gets or sets hypelink.
    /// </summary>
    string? Hyperlink { get; set; }

    /// <summary>
    ///     Gets field.
    /// </summary>
    IField? Field { get; }
}