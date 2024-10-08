﻿// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a text frame.
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
    ///     Gets or sets left margin of text frame in centimeters.
    /// </summary>
    decimal LeftMargin { get; set; }

    /// <summary>
    ///     Gets or sets right margin of text frame in centimeters.
    /// </summary>
    decimal RightMargin { get; set; }

    /// <summary>
    ///     Gets or sets top margin of text frame in centimeters.
    /// </summary>
    decimal TopMargin { get; set; }

    /// <summary>
    ///     Gets or sets bottom margin of text frame in centimeters.
    /// </summary>
    decimal BottomMargin { get; set; }

    /// <summary>
    ///     Gets a value indicating whether the text is wrapped in the shape.
    /// </summary>
    bool TextWrapped { get; }

    /// <summary>
    ///     Gets XPath.
    /// </summary>
    public string SDKXPath { get; }
}