#pragma warning disable IDE0130
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
    ///     Gets Hyperlink. Returns <see langword="null"/> if the portion type doesn't support hyperlink.
    /// </summary>
    IHyperlink? Link { get; }

    /// <summary>
    ///     Gets or sets Text Highlight Color. Returns Color.Transparent if no highlight present.
    /// </summary>
    Color TextHighlightColor { get; set; }

    /// <summary>
    ///     Removes portion from paragraph.
    /// </summary>
    void Remove();
}