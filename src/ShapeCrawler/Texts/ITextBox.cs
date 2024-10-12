// ReSharper disable CheckNamespace
#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a text box.
/// </summary>
public interface ITextBox
{
    /// <summary>
    ///     Gets the collection of paragraphs.
    /// </summary>
    IParagraphs Paragraphs { get; }

    /// <summary>
    ///     Gets or sets text.
    /// </summary>
    string Text { get; set; }

    /// <summary>
    ///     Gets or sets the text vertical alignment.
    /// </summary>
    TextVerticalAlignment VerticalAlignment { get; set; }

    /// <summary>
    ///     Gets or sets the Autofit type.
    /// </summary>
    AutofitType AutofitType { get; set; }

    /// <summary>
    ///     Gets or sets the left margin in centimeters.
    /// </summary>
    decimal LeftMargin { get; set; }

    /// <summary>
    ///     Gets or sets the right margin in centimeters.
    /// </summary>
    decimal RightMargin { get; set; }

    /// <summary>
    ///     Gets or sets the top margin in centimeters.
    /// </summary>
    decimal TopMargin { get; set; }

    /// <summary>
    ///     Gets or sets the bottom margin in centimeters.
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
}