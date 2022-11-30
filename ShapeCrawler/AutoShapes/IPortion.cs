using System;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

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

    /// <summary>
    ///     Gets or sets Text Highlight Color in hexadecimal format. Returns <see langword="null"/> if Text Highlight Color is not present. 
    /// </summary>
    string? TextHighlightColor { get; set; }
}

internal class SCPortion : IPortion
{
    private readonly ResettableLazy<SCFont> font;
    private readonly A.Field? aField;

    internal SCPortion(A.Text aText, SCParagraph paragraph, A.Field aField)
        : this(aText, paragraph)
    {
        this.aField = aField;
    }

    internal SCPortion(A.Text aText, SCParagraph paragraph)
    {
        this.AText = aText;
        this.ParentParagraph = paragraph;
        this.font = new ResettableLazy<SCFont>(() => new SCFont(this.AText, this));
    }

    #region Public Properties

    /// <inheritdoc/>
    public string Text
    {
        get => this.GetText();
        set => this.SetText(value);
    }

    /// <inheritdoc/>
    public IFont Font => this.font.Value;

    public string? Hyperlink
    {
        get => this.GetHyperlink();
        set => this.SetHyperlink(value);
    }

    public IField? Field => this.GetField();

    public string? TextHighlightColor
    {
        get => this.GetTextHighlightColor();
        set => this.SetTextHighlightColor(value!);
    }

    #endregion Public Properties

    internal A.Text AText { get; }

    internal bool IsRemoved { get; set; }

    internal SCParagraph ParentParagraph { get; }

    private IField? GetField()
    {
        if (this.aField is null)
        {
            return null;
        }
        else
        {
            return new SCField(this.aField);
        }
    }

    private string? GetTextHighlightColor()
    {
        var arPr = this.AText.PreviousSibling<A.RunProperties>();
        var aSrgbClr = arPr?.GetFirstChild<A.Highlight>()?.RgbColorModelHex;
        
        return aSrgbClr?.Val;
    }

    private void SetTextHighlightColor(string hex)
    {
        var arPr = this.AText.PreviousSibling<A.RunProperties>() ?? this.AText.Parent!.AddRunProperties();
        
        arPr.AddAHighlight(hex!);
    }
    
    private string GetText()
    {
        var portionText = this.AText.Text;
        if (this.AText.Parent!.NextSibling<A.Break>() != null)
        {
            portionText += Environment.NewLine;
        }

        return portionText;
    }

    private void SetText(string text)
    {
        this.AText.Text = text;
    }

    private string? GetHyperlink()
    {
        var runProperties = this.AText.PreviousSibling<A.RunProperties>();
        if (runProperties == null)
        {
            return null;
        }

        var hyperlink = runProperties.GetFirstChild<A.HyperlinkOnClick>();
        if (hyperlink == null)
        {
            return null;
        }

        var slideObject = (SlideObject)this.ParentParagraph.ParentTextFrame.TextFrameContainer.Shape.SlideObject;
        var typedOpenXmlPart = slideObject.TypedOpenXmlPart;
        var hyperlinkRelationship = (HyperlinkRelationship)typedOpenXmlPart.GetReferenceRelationship(hyperlink.Id!);

        return hyperlinkRelationship.Uri.AbsoluteUri;
    }

    private void SetHyperlink(string? url)
    {
        if (!Uri.IsWellFormedUriString(url, UriKind.Absolute))
        {
            throw new ShapeCrawlerException("Hyperlink is invalid.");
        }

        var runProperties = this.AText.PreviousSibling<A.RunProperties>();
        if (runProperties == null)
        {
            runProperties = new A.RunProperties();
        }

        var hyperlink = runProperties.GetFirstChild<A.HyperlinkOnClick>();
        if (hyperlink == null)
        {
            hyperlink = new A.HyperlinkOnClick();
            runProperties.Append(hyperlink);
        }

        var slideObject = (SlideObject)this.ParentParagraph.ParentTextFrame.TextFrameContainer.Shape.SlideObject;
        var slidePart = slideObject.TypedOpenXmlPart;

        var uri = new Uri(url, UriKind.Absolute);
        var addedHyperlinkRelationship = slidePart.AddHyperlinkRelationship(uri, true);

        hyperlink.Id = addedHyperlinkRelationship.Id;
    }
}