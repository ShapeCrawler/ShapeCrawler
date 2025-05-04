using A = DocumentFormat.OpenXml.Drawing;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a text portion font.
/// </summary>
public interface ITextPortionFont : IFont
{
    /// <summary>
    ///     Gets or sets a font for the Latin characters. Returns <c>null</c> if the Latin font is not present.
    /// </summary>
    string? LatinName { get; set; }
    
    /// <summary>
    ///     Gets or sets a font for the East Asian characters.
    /// </summary>
    string EastAsianName { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether font width is bold.
    /// </summary>
    bool IsBold { get; set; }

    /// <summary>
    ///     Gets or sets a value indicating whether font is italic.
    /// </summary>
    bool IsItalic { get; set; }

    /// <summary>
    ///     Gets or sets the offset effect percentage to make font superscript or subscript.
    /// </summary>
    int OffsetEffect { get; set; }

    /// <summary>
    ///     Gets or sets an underline.
    /// </summary>
    A.TextUnderlineValues Underline { get; set; }

    /// <summary>
    ///     Gets the font color.
    /// </summary>
    IFontColor Color { get; }
}