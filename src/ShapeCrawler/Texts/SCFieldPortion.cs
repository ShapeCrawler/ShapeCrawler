using System;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;

namespace ShapeCrawler.Texts;

internal sealed class SCFieldPortion : IPortion
{
    private readonly ResetableLazy<ITextPortionFont> font;
    private readonly DocumentFormat.OpenXml.Drawing.Field aField;
    private readonly SlideStructure slideStructure;

    internal SCFieldPortion(
        DocumentFormat.OpenXml.Drawing.Field aField, 
        SlideStructure slideStructure, 
        ITextFrameContainer textFrameContainer, 
        SCParagraph paragraph,
        Action onRemoveHandler)
    {
        this.aField = aField;
        this.slideStructure = slideStructure;
        this.AText = aField.GetFirstChild<DocumentFormat.OpenXml.Drawing.Text>() !;
        this.Removed += onRemoveHandler;
        this.font = new ResetableLazy<ITextPortionFont>(() =>
        {
            if (slideStructure is SCSlideLayout &&
                textFrameContainer.SCShape.Placeholder?.Type == SCPlaceholderType.SlideNumber)
            {
                var aListStyle = paragraph.AParagraph.Parent!.GetFirstChild<DocumentFormat.OpenXml.Drawing.ListStyle>() !;
                return new SCLayoutSlideNumberFont(this.AText, this, textFrameContainer, paragraph, aListStyle);    
            }

            return new SCTextPortionFont(this.AText, this, textFrameContainer, paragraph);
        });
    }

    internal event Action? Removed;

    /// <inheritdoc/>
    public string? Text
    {
        get => this.ParseText();
        set => this.SetText(value);
    }

    /// <inheritdoc/>
    public ITextPortionFont Font => this.font.Value;

    public string? Hyperlink
    {
        get => this.GetHyperlink();
        set => this.SetHyperlink(value);
    }

    public IField? Field => this.GetField();

    public SCColor? TextHighlightColor
    {
        get => this.ParseTextHighlightColor();
        set => this.SetTextHighlightColor(value);
    }

    internal DocumentFormat.OpenXml.Drawing.Text AText { get; }

    internal bool IsRemoved { get; set; }
    
    public void Remove()
    {
        this.aField.Remove();
        this.Removed?.Invoke();
    }

    private IField? GetField()
    {
        if (this.aField is null)
        {
            return null;
        }

        return new SCField(this.aField);
    }

    private SCColor ParseTextHighlightColor()
    {
        var arPr = this.AText.PreviousSibling<DocumentFormat.OpenXml.Drawing.RunProperties>();

        // Ensure RgbColorModelHex exists and his value is not null.
        if (arPr?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Highlight>()?.RgbColorModelHex is not DocumentFormat.OpenXml.Drawing.RgbColorModelHex aSrgbClr
            || aSrgbClr.Val is null)
        {
            return SCColor.Transparent;
        }

        // Gets node value.
        // TODO: Check if DocumentFormat.OpenXml.StringValue is necessary.
        var hex = aSrgbClr.Val.ToString() !;

        // Check if color value is valid, we are expecting values as "000000".
        var color = SCColor.FromHex(hex);

        // Calculate alpha value if is defined in highlight node.
        var aAlphaValue = aSrgbClr.GetFirstChild<DocumentFormat.OpenXml.Drawing.Alpha>()?.Val ?? 100000;
        color.Alpha = SCColor.OPACITY / (100000 / aAlphaValue);

        return color;
    }

    private void SetTextHighlightColor(SCColor? color)
    {
        var arPr = this.AText.PreviousSibling<DocumentFormat.OpenXml.Drawing.RunProperties>() ?? this.AText.Parent!.AddRunProperties();

        arPr.AddAHighlight((SCColor)color);
    }

    private string? ParseText()
    {
        var portionText = this.AText?.Text;
        return portionText;
    }

    private void SetText(string? text)
    {
        if (text is null)
        {
            throw new ArgumentNullException(nameof(text));
        }
        
        this.AText.Text = text;
    }

    private string? GetHyperlink()
    {
        var runProperties = this.AText.PreviousSibling<DocumentFormat.OpenXml.Drawing.RunProperties>();
        if (runProperties == null)
        {
            return null;
        }

        var hyperlink = runProperties.GetFirstChild<DocumentFormat.OpenXml.Drawing.HyperlinkOnClick>();
        if (hyperlink == null)
        {
            return null;
        }
        
        var typedOpenXmlPart = this.slideStructure.TypedOpenXmlPart;
        var hyperlinkRelationship = (HyperlinkRelationship)typedOpenXmlPart.GetReferenceRelationship(hyperlink.Id!);

        return hyperlinkRelationship.Uri.ToString();
    }

    private void SetHyperlink(string? url)
    {
        var runProperties = this.AText.PreviousSibling<DocumentFormat.OpenXml.Drawing.RunProperties>();
        if (runProperties == null)
        {
            runProperties = new DocumentFormat.OpenXml.Drawing.RunProperties();
        }

        var hyperlink = runProperties.GetFirstChild<DocumentFormat.OpenXml.Drawing.HyperlinkOnClick>();
        if (hyperlink == null)
        {
            hyperlink = new DocumentFormat.OpenXml.Drawing.HyperlinkOnClick();
            runProperties.Append(hyperlink);
        }
        
        var slidePart = this.slideStructure.TypedOpenXmlPart;

        var uri = new Uri(url!, UriKind.RelativeOrAbsolute);
        var addedHyperlinkRelationship = slidePart.AddHyperlinkRelationship(uri, true);

        hyperlink.Id = addedHyperlinkRelationship.Id;
    }
}