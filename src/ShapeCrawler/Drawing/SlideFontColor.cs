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
    private bool initialized;
    private string? hexColor;
    private SCColorType colorType;
    private readonly SlidePart sdkSlidePart;

    internal SlideFontColor(SlidePart sdkSlidePart, A.Text aText)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.aText = aText;
    }

    public SCColorType ColorType => this.ParseColorType();

    public string ColorHex => this.ParseHexColor();

    public void SetColorByHex(string hex)
    {
        var aTextContainer = this.aText.Parent!;
        var aRunProperties = aTextContainer.GetFirstChild<A.RunProperties>() ?? aTextContainer.AddRunProperties();

        var aSolidFill = aRunProperties.GetASolidFill();
        aSolidFill?.Remove();

        // All hex values are expected to be without hashtag.
        hex = hex.StartsWith("#", System.StringComparison.Ordinal) ? hex.Substring(1) : hex; // to skip '#'
        var rgbColorModelHex = new A.RgbColorModelHex { Val = hex };
        aSolidFill = new A.SolidFill();
        aSolidFill.Append(rgbColorModelHex);
        aRunProperties.Append(aSolidFill);
    }

    private SCColorType ParseColorType()
    {
        if (!this.initialized)
        {
            this.InitializeColor();
        }

        return this.colorType;
    }

    private string ParseHexColor()
    {
        if (!this.initialized)
        {
            this.InitializeColor();
        }

        return this.hexColor!;
    }

    private void InitializeColor()
    {
        var pSlideMaster = this.sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster;
        this.initialized = true;
        var aSolidFill = this.aText.Parent!.GetFirstChild<A.RunProperties>()?.GetASolidFill();
        if (aSolidFill != null)
        {
            var typeAndHex = HexParser.FromSolidFill(aSolidFill, pSlideMaster);
            this.colorType = typeAndHex.Item1;
            this.hexColor = typeAndHex.Item2;
        }
        else
        {
            if (this.TryFromTextBody())
            {
                return;
            }

            if (this.TryFromShapeFontReference())
            {
                return;
            }

            var aParagraph = new SDKOpenXmlElementWrap(this.aText).FirstAncestor<A.Paragraph>();
            var paraLevel = new SDKAParagraphWrap(aParagraph).IndentLevel();

            var pSlideMasterWrap =
                new SDKPSlideMasterWrap(pSlideMaster);
            ParagraphLevelFont? masterBodyStyleParaLevelFont = pSlideMasterWrap.BodyStyleFontOrNull(paraLevel);
            if (this.TryFromFontData(masterBodyStyleParaLevelFont))
            {
                return;
            }

            // Presentation level
            var sdkSlidePartWrap = new SDKSlidePartWrap(this.sdkSlidePart);
            ParagraphLevelFont? presParaLevelFont = sdkSlidePartWrap.PresentationFontOrThemeFontOrNull(paraLevel);
            string hexColor;
            if (presParaLevelFont.HasValue)
            {
                hexColor = sdkSlidePartWrap.ThemeHexColor(presParaLevelFont.Value.ASchemeColor!.Val!);
                this.colorType = SCColorType.Theme;
                this.hexColor = hexColor;
                return;
            }

            // Get default
            hexColor = sdkSlidePartWrap.ThemeHexColor(A.SchemeColorValues.Text1);
            this.colorType = SCColorType.Theme;
            this.hexColor = hexColor;
        }
    }

    private bool TryFromTextBody()
    {
        var aParagraph = new SDKOpenXmlElementWrap(this.aText).FirstAncestor<A.Paragraph>();
        var paraLevel = new SDKAParagraphWrap(aParagraph).IndentLevel();
        var pTextBody = new SDKOpenXmlElementWrap(aParagraph).FirstAncestor<P.TextBody>();
        var textBodyStyleFont = new ParagraphLevelFonts(pTextBody.GetFirstChild<A.ListStyle>()!).FontOrNull(paraLevel);

        if (textBodyStyleFont.HasValue)
        {
            return this.TryFromFontData(textBodyStyleFont);
        }

        return false;
    }

    private bool TryFromShapeFontReference()
    {
        P.Shape parentPShape = new SDKOpenXmlElementWrap(this.aText).FirstAncestor<P.Shape>();
        if (parentPShape.ShapeStyle == null)
        {
            return false;
        }

        var aFontReference = parentPShape.ShapeStyle.FontReference!;
        var fontReferenceFontData = new ParagraphLevelFont
        {
            ARgbColorModelHex = aFontReference.RgbColorModelHex,
            ASchemeColor = aFontReference.SchemeColor,
            APresetColor = aFontReference.PresetColor
        };

        return this.TryFromFontData(fontReferenceFontData);
    }

    private bool TryFromFontData(ParagraphLevelFont? paragraphLevelFont)
    {
        if (!paragraphLevelFont.HasValue)
        {
            return false;
        }

        string colorHexVariant;
        if (paragraphLevelFont.Value.ARgbColorModelHex != null)
        {
            colorHexVariant = paragraphLevelFont.Value.ARgbColorModelHex.Val!;
            this.colorType = SCColorType.RGB;
            this.hexColor = colorHexVariant;
            return true;
        }

        if (paragraphLevelFont.Value.ASchemeColor != null)
        {
            var sdkSlidePartWrap = new SDKSlidePartWrap(this.sdkSlidePart);
            colorHexVariant = sdkSlidePartWrap.ThemeHexColor(paragraphLevelFont.Value.ASchemeColor.Val!);
            this.colorType = SCColorType.Theme;
            this.hexColor = colorHexVariant;
            return true;
        }

        if (paragraphLevelFont.Value.ASystemColor != null)
        {
            colorHexVariant = paragraphLevelFont.Value.ASystemColor.LastColor!;
            this.colorType = SCColorType.Standard;
            this.hexColor = colorHexVariant;
            return true;
        }

        if (paragraphLevelFont.Value.APresetColor != null)
        {
            this.colorType = SCColorType.Preset;
            var coloName = paragraphLevelFont.Value.APresetColor.Val!.Value.ToString();
            this.hexColor = SCColorTranslator.HexFromName(coloName);
            return true;
        }

        return false;
    }
}