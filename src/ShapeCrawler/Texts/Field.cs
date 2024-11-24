using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Fonts;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Texts;

internal sealed class Field : IParagraphPortion
{
    private readonly OpenXmlPart sdkTypedOpenXmlPart;
    private readonly ResetableLazy<ITextPortionFont> font;
    private readonly Lazy<Hyperlink> hyperlink;
    private readonly A.Field aField;
    private readonly PortionText portionText;
    private readonly A.Text? aText;

    internal Field(OpenXmlPart sdkTypedOpenXmlPart, A.Field aField)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.aText = aField.GetFirstChild<A.Text>();
        this.aField = aField;

        this.font = new ResetableLazy<ITextPortionFont>(() =>
        {
            var textPortionSize = new PortionFontSize(sdkTypedOpenXmlPart, this.aText!);
            return new TextPortionFont(sdkTypedOpenXmlPart, this.aText!, textPortionSize);
        });

        this.portionText = new PortionText(this.aField);
        this.hyperlink = new Lazy<Hyperlink>(() => new Hyperlink(this.aField.RunProperties!));
    }

    /// <inheritdoc/>
    public string? Text
    {
        get => this.portionText.Text();
        set => this.portionText.Update(value!);
    }

    /// <inheritdoc/>
    public ITextPortionFont Font => this.font.Value;

    public IHyperlink Link => this.hyperlink.Value;

    public Color TextHighlightColor
    {
        get => this.ParseTextHighlight();
        set => this.UpdateTextHighlight(value);
    }

    public void Remove() => this.aField.Remove();

    private Color ParseTextHighlight()
    {
        var arPr = this.aText!.PreviousSibling<A.RunProperties>();

        // Ensure RgbColorModelHex exists and his value is not null.
        if (arPr?.GetFirstChild<A.Highlight>()?.RgbColorModelHex is not A.RgbColorModelHex aSrgbClr
            || aSrgbClr.Val is null)
        {
            return Color.Transparent;
        }

        // Gets node value.
        // TODO: Check if DocumentFormat.OpenXml.StringValue is necessary.
        var hex = aSrgbClr.Val.ToString() !;

        // Check if color value is valid, we are expecting values as "000000".
        var color = Color.FromHex(hex);

        // Calculate alpha value if is defined in highlight node.
        var aAlphaValue = aSrgbClr.GetFirstChild<A.Alpha>()?.Val ?? 100000;
        color.Alpha = Color.Opacity / (100000 / aAlphaValue);

        return color;
    }

    private void UpdateTextHighlight(Color color)
    {        
        var arPr = this.aText!.PreviousSibling<A.RunProperties>() ?? this.aText.Parent!.AddRunProperties();

        arPr.AddAHighlight(color);
    }

    private string? GetHyperlink()
    {
        var runProperties = this.aText!.PreviousSibling<A.RunProperties>();
        if (runProperties == null)
        {
            return null;
        }

        var hyperlink = runProperties.GetFirstChild<A.HyperlinkOnClick>();
        if (hyperlink == null)
        {
            return null;
        }

        var hyperlinkRelationship = (HyperlinkRelationship)this.sdkTypedOpenXmlPart.GetReferenceRelationship(hyperlink.Id!);

        return hyperlinkRelationship.Uri.ToString();
    }

    private void SetHyperlink(string? url)
    {
        if (url is null)
        {
            throw new SCException("URL is null.");
        }

        var runProperties = this.aText!.PreviousSibling<A.RunProperties>() ?? new A.RunProperties();

        var hyperlink = runProperties.GetFirstChild<A.HyperlinkOnClick>();
        if (hyperlink == null)
        {
            hyperlink = new A.HyperlinkOnClick();
            runProperties.Append(hyperlink);
        }

        if (url.StartsWith("slide://"))
        {
            // Handle inner slide hyperlink
            var slideNumber = int.Parse(url.Substring(8));
            var presentation = ((PresentationDocument)this.sdkTypedOpenXmlPart.OpenXmlPackage).PresentationPart!;
            var slideId = presentation.Presentation.SlideIdList!.ChildElements
                .OfType<P.SlideId>()
                .ElementAtOrDefault(slideNumber - 1);

            if (slideId == null)
            {
                throw new SCException($"Invalid slide number: {slideNumber}");
            }

            // Get the target slide part
            var targetSlidePart = presentation.GetPartById(slideId.RelationshipId!) as SlidePart;
            if (targetSlidePart == null)
            {
                throw new SCException($"Could not find slide part for slide {slideNumber}");
            }

            // Add relationship from current slide to target slide
            var currentSlidePart = this.sdkTypedOpenXmlPart as SlidePart;
            if (currentSlidePart == null)
            {
                throw new SCException("Current part is not a slide part");
            }

            // Add or reuse relationship to target slide
            var relationship = currentSlidePart.GetIdOfPart(targetSlidePart);
            if (string.IsNullOrEmpty(relationship))
            {
                // relationship = currentSlidePart.AddPart(targetSlidePart);
            }

            hyperlink.Id = relationship;
            hyperlink.Action = "ppaction://hlinksldjump";
        }
        else
        {
            // Handle regular URL or file hyperlink
            var uri = new Uri(url, UriKind.RelativeOrAbsolute);
            var addedHyperlinkRelationship = this.sdkTypedOpenXmlPart.AddHyperlinkRelationship(uri, true);
            hyperlink.Id = addedHyperlinkRelationship.Id;
            hyperlink.Action = null;
        }
    }
}