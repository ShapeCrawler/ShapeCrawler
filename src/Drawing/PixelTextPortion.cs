using System.Linq;

namespace ShapeCrawler.Drawing;

/// <summary>
///     Represents a measured text portion.
/// </summary>
internal readonly struct PixelTextPortion(string text, ITextPortionFont? font, float width)
{
    /// <summary>
    ///     Gets text content.
    /// </summary>
    internal string Text => text;

    /// <summary>
    ///     Gets text font.
    /// </summary>
    internal ITextPortionFont? Font => font;

    /// <summary>
    ///     Gets text width in pixels.
    /// </summary>
    internal float Width => width;

    /// <summary>
    ///     Gets a value indicating whether the text portion contains only whitespace.
    /// </summary>
    internal bool IsWhitespace => string.IsNullOrEmpty(this.Text) || this.Text.All(char.IsWhiteSpace);
}