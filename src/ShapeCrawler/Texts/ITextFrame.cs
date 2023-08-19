// ReSharper disable CheckNamespace

using ShapeCrawler.Texts;

namespace ShapeCrawler;

/// <summary>
///     Represents a text frame.
/// </summary>
public interface ITextFrame
{
    /// <summary>
    ///     Gets collection of paragraphs.
    /// </summary>
    IParagraphCollection Paragraphs { get; }

    /// <summary>
    ///     Gets or sets text.
    /// </summary>
    string Text { get; set; }

    /// <summary>
    ///     Gets or sets Autofit type.
    /// </summary>
    SCAutofitType AutofitType { get; set; }

    /// <summary>
    ///     Gets or sets left margin of text frame in centimeters.
    /// </summary>
    double LeftMargin { get; set; }

    /// <summary>
    ///     Gets or sets right margin of text frame in centimeters.
    /// </summary>
    double RightMargin { get; set; }

    /// <summary>
    ///     Gets or sets top margin of text frame in centimeters.
    /// </summary>
    double TopMargin { get; set; }

    /// <summary>
    ///     Gets or sets bottom margin of text frame in centimeters.
    /// </summary>
    double BottomMargin { get; set; }

    /// <summary>
    ///     Gets a value indicating whether text is wrapped in shape.
    /// </summary>
    bool TextWrapped { get; }
}