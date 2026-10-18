// ReSharper disable CheckNamespace
// ReSharper disable InconsistentNaming
#pragma warning disable IDE0130
namespace ShapeCrawler;

using ShapeCrawler.Texts;
#pragma warning restore IDE0130

/// <summary>
///     Represents a PowerPoint Text Box property (Format Shape → Shape Options → Size &amp; Properties → Text Box).
/// </summary>
public interface ITextBox
{
    /// <summary>
    ///     Gets the collection of paragraphs.
    /// </summary>
    IParagraphCollection Paragraphs { get; }

    /// <summary>
    ///     Gets text.
    /// </summary>
    string Text { get; }

    /// <summary>
    ///     Gets or sets the text vertical alignment.
    /// </summary>
    TextVerticalAlignment VerticalAlignment { get; set; }

    /// <summary>
    ///     Gets or sets the Autofit type.
    /// </summary>
    AutofitType AutofitType { get; set; }

    /// <summary>
    ///     Gets or sets the TextDirection.
    /// </summary>
    TextDirection TextDirection { get; set; }

    /// <summary>
    ///     Gets or sets the left margin in points.
    /// </summary>
    decimal LeftMargin { get; set; }

    /// <summary>
    ///     Gets or sets the right margin in points.
    /// </summary>
    decimal RightMargin { get; set; }

    /// <summary>
    ///     Gets or sets the top margin in points.
    /// </summary>
    decimal TopMargin { get; set; }

    /// <summary>
    ///     Gets or sets the bottom margin in points.
    /// </summary>
    decimal BottomMargin { get; set; }

    /// <summary>
    ///     Gets a value indicating whether the text is wrapped in the shape.
    /// </summary>
    bool TextWrapped { get; }

    /// <summary>
    ///     Gets XPath.
    /// </summary>
    public string SdkXPath { get; }

    /// <summary>
    ///     Sets text.
    /// </summary>
    public void SetText(string text);

    /// <summary>
    ///     Replaces literal, non-overlapping matches, including matches split across text runs.
    ///     Matching is ordinal and case-sensitive unless ignoreCase is true.
    ///     Matches do not cross paragraphs, line breaks, or fields. Replacement text inherits
    ///     the first matched run's formatting; other text, font sizes, and geometry are retained.
    ///     Newlines in the replacement become soft line breaks. Returns the number of replacements.
    /// </summary>
    int ReplaceText(string oldText, string newText, bool ignoreCase = false);

    /// <summary>
    ///     Sets Markdown text.
    /// </summary>
    public void SetMarkdownText(string text);
}
