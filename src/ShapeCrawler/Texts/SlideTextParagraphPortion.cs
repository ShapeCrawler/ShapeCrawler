using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Fonts;
using ShapeCrawler.Services;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal sealed class SlideTextParagraphPortion : IParagraphPortion
{
    private readonly ResetableLazy<SlideTextPortionFont> font;
    private readonly SlidePart sdkSlidePart;
    private readonly A.Run aRun;

    internal SlideTextParagraphPortion(SlidePart sdkSlidePart, A.Run aRun)
    {
        this.AText = aRun.Text!;
        this.sdkSlidePart = sdkSlidePart;
        this.aRun = aRun;
        var textPortionSize = new PortionFontSize(this.AText);
        this.font = new ResetableLazy<SlideTextPortionFont>(() => new SlideTextPortionFont(sdkSlidePart,this.AText, textPortionSize));
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

    public SCColor? TextHighlightColor
    {
        get => this.ParseTextHighlight();
        set => this.UpdateTextHighlight(value);
    }

    internal A.Text AText { get; }

    internal bool IsRemoved { get; set; }
    
    public void Remove()
    {
        this.aRun.Remove();
        this.Removed?.Invoke();
    }

    private SCColor ParseTextHighlight()
    {
        var arPr = this.AText.PreviousSibling<A.RunProperties>();

        // Ensure RgbColorModelHex exists and his value is not null.
        if (arPr?.GetFirstChild<A.Highlight>()?.RgbColorModelHex is not A.RgbColorModelHex aSrgbClr
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
        var aAlphaValue = aSrgbClr.GetFirstChild<A.Alpha>()?.Val ?? 100000;
        color.Alpha = SCColor.OPACITY / (100000 / aAlphaValue);

        return color;
    }

    private void UpdateTextHighlight(SCColor? color)
    {
        var arPr = this.AText.PreviousSibling<A.RunProperties>() ?? this.AText.Parent!.AddRunProperties();

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
        
        var hyperlinkRelationship = (HyperlinkRelationship)this.sdkSlidePart.GetReferenceRelationship(hyperlink.Id!);

        return hyperlinkRelationship.Uri.ToString();
    }

    private void SetHyperlink(string? url)
    {
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
        
        var uri = new Uri(url!, UriKind.RelativeOrAbsolute);
        var addedHyperlinkRelationship = this.sdkSlidePart.AddHyperlinkRelationship(uri, true);

        hyperlink.Id = addedHyperlinkRelationship.Id;
    }
}