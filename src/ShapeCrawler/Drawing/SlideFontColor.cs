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

        // TryFromTextBody()
        var aParagraph1 = new SDKOpenXmlElementWrap(this.aText).FirstAncestor<A.Paragraph>();
        var paraLevel1 = new SDKAParagraphWrap(aParagraph1).IndentLevel();
        var pTextBody = new SDKOpenXmlElementWrap(aParagraph1).FirstAncestor<P.TextBody>();
        var textBodyStyleFont = new ParagraphLevelFonts(pTextBody.GetFirstChild<A.ListStyle>()!).FontOrNull(paraLevel1);
        if (textBodyStyleFont.HasValue)
        {
            if(this.TryFromParaLevelFont(textBodyStyleFont, out var response))
            {
                return response.colorHex!;
            }
        }

        // TryFromShapeFontReference()
        var sdkPShape = new SDKOpenXmlElementWrap(this.aText).FirstAncestor<P.Shape>();
        var sdkPShapeWrap = new SDKPShapeWrap(sdkPShape);
        string? shapeFontColorHex = sdkPShapeWrap.FontColorHexOrNull();
        if (shapeFontColorHex != null)
        {
            return shapeFontColorHex;
        }
        
        var aParagraph = new SDKOpenXmlElementWrap(this.aText).FirstAncestor<A.Paragraph>();
        var paraLevel = new SDKAParagraphWrap(aParagraph).IndentLevel();

        var pSlideMasterWrap =
            new SDKPSlideMasterWrap(sdkPSlideMaster);
        ParagraphLevelFont? masterBodyStyleParaLevelFont = pSlideMasterWrap.BodyStyleFontOrNull(paraLevel);
        if (this.TryFromParaLevelFont(masterBodyStyleParaLevelFont))
        {
            return;
        }

        // Presentation level
        var sdkSlidePartWrap = new SDKSlidePartWrap(this.sdkSlidePart);
        ParagraphLevelFont? presParaLevelFont = sdkSlidePartWrap.PresentationFontOrThemeFontOrNull(paraLevel);
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

    private bool TryFromTextBody()
    {
        var aParagraph = new SDKOpenXmlElementWrap(this.aText).FirstAncestor<A.Paragraph>();
        var paraLevel = new SDKAParagraphWrap(aParagraph).IndentLevel();
        var pTextBody = new SDKOpenXmlElementWrap(aParagraph).FirstAncestor<P.TextBody>();
        var textBodyStyleFont = new ParagraphLevelFonts(pTextBody.GetFirstChild<A.ListStyle>()!).FontOrNull(paraLevel);

        if (textBodyStyleFont.HasValue)
        {
            return this.TryFromParaLevelFont(textBodyStyleFont);
        }

        return false;
    }

    private bool TryFromShapeFontReference()
    {
        var parentSDKPShape = new SDKOpenXmlElementWrap(this.aText).FirstAncestor<P.Shape>();
        if (parentSDKPShape.ShapeStyle == null)
        {
            return false;
        }

        var aFontReference = parentSDKPShape.ShapeStyle.FontReference!;
        var fontReferenceFontData = new ParagraphLevelFont
        {
            ARgbColorModelHex = aFontReference.RgbColorModelHex,
            ASchemeColor = aFontReference.SchemeColor,
            APresetColor = aFontReference.PresetColor
        };

        return this.TryFromParaLevelFont(fontReferenceFontData);
    }

    private bool TryFromParaLevelFont(ParagraphLevelFont? paraLevelFont, out (SCColorType colorType, string? colorHex) response)
    {
        if (!paraLevelFont.HasValue)
        {
            response = (SCColorType.NotDefined, null);
            return false;
        }

        string colorHexVariant;
        if (paraLevelFont.Value.ARgbColorModelHex != null)
        {
            colorHexVariant = paraLevelFont.Value.ARgbColorModelHex.Val!;
            response = (SCColorType.RGB, colorHexVariant);
            return true;
        }

        if (paraLevelFont.Value.ASchemeColor != null)
        {
            var sdkSlidePartWrap = new SDKSlidePartWrap(this.sdkSlidePart);
            colorHexVariant = sdkSlidePartWrap.ThemeColorHex(paraLevelFont.Value.ASchemeColor.Val!);
            response = (SCColorType.Theme, colorHexVariant);
            return true;
        }

        if (paraLevelFont.Value.ASystemColor != null)
        {
            colorHexVariant = paraLevelFont.Value.ASystemColor.LastColor!;
            response = (SCColorType.Standard, colorHexVariant);
            return true;
        }

        if (paraLevelFont.Value.APresetColor != null)
        {
            var coloName = paraLevelFont.Value.APresetColor.Val!.Value.ToString();
            response = (SCColorType.Preset, ColorTranslator.HexFromName(coloName));
            return true;
        }

        response = (SCColorType.NotDefined, null);
        return false;
    }
}