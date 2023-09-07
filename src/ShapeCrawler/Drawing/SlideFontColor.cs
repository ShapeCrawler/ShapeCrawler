using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using ShapeCrawler.Fonts;
using ShapeCrawler.Wrappers;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal sealed class SlideFontColor : IFontColor
{
    private readonly A.Text aText;
    private readonly SlidePart sdkSlidePart;

    internal SlideFontColor(SlidePart sdkSlidePart, A.Text aText)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.aText = aText;
    }

    public SCColorType ColorType => this.ParseColorType();

    public string ColorHex => this.ParseColorHex();

    public void SetColorByHex(string hex)
    {
        var aTextContainer = this.aText.Parent!;
        var aRunProperties = aTextContainer.GetFirstChild<A.RunProperties>() ?? aTextContainer.AddRunProperties();

        var aSolidFill = aRunProperties.SDKASolidFill();
        aSolidFill?.Remove();

        // All hex values are expected to be without hashtag.
        hex = hex.StartsWith("#", System.StringComparison.Ordinal) ? hex.Substring(1) : hex; // to skip '#'
        var rgbColorModelHex = new A.RgbColorModelHex { Val = hex };
        aSolidFill = new A.SolidFill();
        aSolidFill.Append(rgbColorModelHex);
        aRunProperties.Append(aSolidFill);
    }

    private string ParseColorHex()
    {
        var sdkPSlideMaster = this.sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster;
        var sdkASolidFill = this.aText.Parent!.GetFirstChild<A.RunProperties>()?.SDKASolidFill();
        if (sdkASolidFill != null)
        {
            var typeAndColor = HexParser.FromSolidFill(sdkASolidFill, sdkPSlideMaster);
            return typeAndColor.Item2!;
        }

        // From TextBody
        var aParagraph = new SdkOpenXmlElement(this.aText).FirstAncestor<A.Paragraph>();
        var indentLevel = new SdkAParagraphWrap(aParagraph).IndentLevel();
        var pTextBody = new SdkOpenXmlElement(aParagraph).FirstAncestor<P.TextBody>();
        var textBodyStyleFont = new IndentFonts(pTextBody.GetFirstChild<A.ListStyle>()!).FontOrNull(indentLevel);
        if (textBodyStyleFont.HasValue)
        {
            if (this.TryFromIndentFont(textBodyStyleFont, out var textBodyColor))
            {
                return textBodyColor.colorHex!;
            }
        }

        // From Shape
        var pShape = new SdkOpenXmlElement(this.aText).FirstAncestor<P.Shape>();
        var sdkSlidePShapeWrap = new SdkSlidePShape(new SdkSlidePart(this.sdkSlidePart), pShape);
        string? shapeFontColorHex = sdkSlidePShapeWrap.FontColorHexOrNull();
        if (shapeFontColorHex != null)
        {
            return shapeFontColorHex;
        }
        
        // From Referenced Layout Shape
        var sdkSlidePartWrap = new SdkSlidePart(this.sdkSlidePart);
        var refShapeFontColorHex = sdkSlidePartWrap.ReferencedShapeColorOrNull(pShape, indentLevel);
        if (refShapeFontColorHex != null)
        {
            return refShapeFontColorHex;
        }
        
        // From Common Placeholder

        var pSlideMasterWrap =
            new SdkPSlideMasterWrap(sdkPSlideMaster);
        var masterIndentFont = pSlideMasterWrap.BodyStyleFontOrNull(indentLevel);
        if (this.TryFromIndentFont(masterIndentFont, out var masterColor))
        {
            return masterColor.colorHex!;
        }

        // Presentation level
        IndentFont? presParaLevelFont = sdkSlidePartWrap.PresentationFontOrThemeFontOrNull(indentLevel);
        string colorHex;
        if (presParaLevelFont.HasValue)
        {
            colorHex = sdkSlidePartWrap.ThemeColorHex(presParaLevelFont.Value.ASchemeColor!.Val!);
            return colorHex;
        }

        // Get default
        colorHex = sdkSlidePartWrap.ThemeColorHex(A.SchemeColorValues.Text1);
        return colorHex;
    }

    private SCColorType ParseColorType()
    {
        var sdkPSlideMaster = this.sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster;
        var sdkASolidFill = this.aText.Parent!.GetFirstChild<A.RunProperties>()?.SDKASolidFill();
        if (sdkASolidFill != null)
        {
            var typeAndColor = HexParser.FromSolidFill(sdkASolidFill, sdkPSlideMaster);
            return typeAndColor.Item1;
        }

        // TryFromTextBody()
        var aParagraph = new SdkOpenXmlElement(this.aText).FirstAncestor<A.Paragraph>();
        var paraLevel = new SdkAParagraphWrap(aParagraph).IndentLevel();
        var pTextBody = new SdkOpenXmlElement(aParagraph).FirstAncestor<P.TextBody>();
        var textBodyStyleFont = new IndentFonts(pTextBody.GetFirstChild<A.ListStyle>()!).FontOrNull(paraLevel);
        if (textBodyStyleFont.HasValue)
        {
            if (this.TryFromIndentFont(textBodyStyleFont, out var textBodyColor))
            {
                return textBodyColor.colorType;
            }
        }

        return SCColorType.NotDefined;
    }

    private bool TryFromIndentFont(
        IndentFont? indentFont,
        out (SCColorType colorType, string? colorHex) response)
    {
        if (!indentFont.HasValue)
        {
            response = (SCColorType.NotDefined, null);
            return false;
        }

        string colorHexVariant;
        if (indentFont.Value.ARgbColorModelHex != null)
        {
            colorHexVariant = indentFont.Value.ARgbColorModelHex.Val!;
            response = (SCColorType.RGB, colorHexVariant);
            return true;
        }

        if (indentFont.Value.ASchemeColor != null)
        {
            var sdkSlidePartWrap = new SdkSlidePart(this.sdkSlidePart);
            colorHexVariant = sdkSlidePartWrap.ThemeColorHex(indentFont.Value.ASchemeColor.Val!);
            response = (SCColorType.Theme, colorHexVariant);
            return true;
        }

        if (indentFont.Value.ASystemColor != null)
        {
            colorHexVariant = indentFont.Value.ASystemColor.LastColor!;
            response = (SCColorType.Standard, colorHexVariant);
            return true;
        }

        if (indentFont.Value.APresetColor != null)
        {
            var coloName = indentFont.Value.APresetColor.Val!.Value.ToString();
            response = (SCColorType.Preset, ColorTranslator.HexFromName(coloName));
            return true;
        }

        response = (SCColorType.NotDefined, null);
        return false;
    }
}