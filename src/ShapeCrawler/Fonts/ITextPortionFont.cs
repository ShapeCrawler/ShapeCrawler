using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents font of a text portion.
/// </summary>
public interface ITextPortionFont : IFont
{
    /// <summary>
    ///     Gets or sets font for the Latin characters. Returns <see langword="null"/> if the Latin font is not present.
    /// </summary>
    string? LatinName { get; set; }
    
    /// <summary>
    ///     Gets or sets font for the East Asian characters.
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
    ///     Gets font color formatting.
    /// </summary>
    IFontColor Color { get; }
}