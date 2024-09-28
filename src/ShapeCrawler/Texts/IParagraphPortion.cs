// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a portion of a paragraph.
/// </summary>
public interface IParagraphPortion
{
    /// <summary>
    ///     Gets or sets text.
    /// </summary>
    string? Text { get; set; }

    /// <summary>
    ///     Gets font.
    /// </summary>
    ITextPortionFont? Font { get; }

    /// <summary>
    ///     Gets or sets hypelink.
    /// </summary>
    string? Hyperlink { get; set; }

    /// <summary>
    ///     Gets or sets Text Highlight Color. Returns Color.Transparent if no highlight present.
    /// </summary>
    Color TextHighlightColor { get; set; }

    /// <summary>
    ///     Removes portion from paragraph.
    /// </summary>
    void Remove();
}