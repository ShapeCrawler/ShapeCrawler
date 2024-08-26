using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Colors;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Fonts;
using ShapeCrawler.ShapeCollection;
using ShapeCrawler.SlideMasters;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

internal sealed class FontColor : IFontColor
{
    private readonly OpenXmlPart sdkTypedOpenXmlPart;
    private readonly A.Text aText;

    internal FontColor(OpenXmlPart sdkTypedOpenXmlPart, A.Text aText)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.aText = aText;
    }

    #region Public APIs

    public ColorType Type
    {
        get
        {
            var aSolidFill = this.aText.Parent!.GetFirstChild<A.RunProperties>()?.SDKASolidFill();
            if (aSolidFill != null)
            {
                var pSlideMaster = this.sdkTypedOpenXmlPart switch
                {
                    SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster,
                    SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideMasterPart!.SlideMaster,
                    _ => ((SlideMasterPart)this.sdkTypedOpenXmlPart).SlideMaster
                };
                var typeAndColor = HexParser.FromSolidFill(aSolidFill, pSlideMaster);
                return typeAndColor.Item1;
            }

            // TryFromTextBody()
            var aParagraph = this.aText.Ancestors<A.Paragraph>().First();
            var indentLevel = new WrappedAParagraph(aParagraph).IndentLevel();
            var pTextBody = aParagraph.Ancestors<P.TextBody>().First();
            var aListStyle = pTextBody.GetFirstChild<A.ListStyle>() !;
            var textBodyStyleFont = new IndentFonts(aListStyle).FontOrNull(indentLevel);
            if (textBodyStyleFont.HasValue)
            {
                if (this.TryFromIndentFont(textBodyStyleFont, out var textBodyColor))
                {
                    return textBodyColor.colorType;
                }
            }

            // From Shape
            var shapeColor = new ShapeColor(this.sdkTypedOpenXmlPart, this.aText);
            var type = shapeColor.TypeOrNull();
            if (type.HasValue)
            {
                return (ColorType)type;
            }

            // From Referenced Shape
            if (this.sdkTypedOpenXmlPart is not SlideMasterPart)
            {
                var refShapeColorType = new ReferencedIndentLevel(this.sdkTypedOpenXmlPart, this.aText).ColorTypeOrNull();
                if (refShapeColorType.HasValue)
                {
                    return (ColorType)refShapeColorType;
                }
            }

            return ColorType.NotDefined;
        }
    }

    public string Hex
    {
        get
        {
            var pSlideMaster = this.sdkTypedOpenXmlPart switch
            {
                SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster,
                SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideMasterPart!.SlideMaster,
                _ => ((SlideMasterPart)this.sdkTypedOpenXmlPart).SlideMaster
            };
            var aSolidFill = this.aText.Parent!.GetFirstChild<A.RunProperties>()?.SDKASolidFill();
            if (aSolidFill != null)
            {
                var typeAndColor = HexParser.FromSolidFill(aSolidFill, pSlideMaster);
                return typeAndColor.Item2!;
            }

            // From TextBody
            var aParagraph = this.aText.Ancestors<A.Paragraph>().First();
            var indentLevel = new WrappedAParagraph(aParagraph).IndentLevel();
            var pTextBody = aParagraph.Ancestors<P.TextBody>().First();
            var textBodyStyleFont = new IndentFonts(pTextBody.GetFirstChild<A.ListStyle>()
                !).FontOrNull(indentLevel);
            if (textBodyStyleFont.HasValue)
            {
                if (this.TryFromIndentFont(textBodyStyleFont, out var textBodyColor))
                {
                    return textBodyColor.colorHex!;
                }
            }

            // From Shape
            var pShape = this.aText.Ancestors<P.Shape>().First();
            var sdkSlidePShapeWrap = new ShapeColor(new PresentationColor(this.sdkTypedOpenXmlPart), pShape);
            string? shapeFontColorHex = sdkSlidePShapeWrap.HexOrNull();
            if (shapeFontColorHex != null)
            {
                return shapeFontColorHex;
            }

            // From Referenced Shape
            if (this.sdkTypedOpenXmlPart is not SlideMasterPart)
            {
                var refShapeFontColorHex = new ReferencedIndentLevel(this.sdkTypedOpenXmlPart, this.aText).ColorHexOrNull();
                if (refShapeFontColorHex != null)
                {
                    return refShapeFontColorHex;
                }
            }

            // From Common Placeholder
            var pSlideMasterWrap =
                new WrappedPSlideMaster(pSlideMaster);
            var masterIndentFont = pSlideMasterWrap.BodyStyleFontOrNull(indentLevel);
            if (this.TryFromIndentFont(masterIndentFont, out var masterColor))
            {
                return masterColor.colorHex!;
            }

            // Presentation level
            var presColor = new PresentationColor(this.sdkTypedOpenXmlPart);
            var presParaLevelFont = presColor.PresentationFontOrThemeFontOrNull(indentLevel);
            string colorHex;
            if (presParaLevelFont.HasValue)
            {
                colorHex = presColor.ThemeColorHex(presParaLevelFont.Value.ASchemeColor!.Val!);
                return colorHex;
            }

            // Get default
            colorHex = presColor.ThemeColorHex(A.SchemeColorValues.Text1);
            return colorHex;
        }
    }

    public void Update(string hex)
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

    #endregion Public APIs

    private bool TryFromIndentFont(
        IndentFont? indentFont,
        out (ColorType colorType, string? colorHex) response)
    {
        if (!indentFont.HasValue)
        {
            response = (ColorType.NotDefined, null);
            return false;
        }

        string colorHexVariant;
        if (indentFont.Value.ARgbColorModelHex != null)
        {
            colorHexVariant = indentFont.Value.ARgbColorModelHex.Val!;
            response = (ColorType.RGB, colorHexVariant);
            return true;
        }

        if (indentFont.Value.ASchemeColor != null)
        {
            var sdkSlidePartWrap = new PresentationColor(this.sdkTypedOpenXmlPart);
            colorHexVariant = sdkSlidePartWrap.ThemeColorHex(indentFont.Value.ASchemeColor.Val!);
            response = (ColorType.Theme, colorHexVariant);
            return true;
        }

        if (indentFont.Value.ASystemColor != null)
        {
            colorHexVariant = indentFont.Value.ASystemColor.LastColor!;
            response = (ColorType.Standard, colorHexVariant);
            return true;
        }

        if (indentFont.Value.APresetColor != null)
        {
            var coloName = indentFont.Value.APresetColor.Val!.Value.ToString();
            response = (ColorType.Preset, ColorTranslator.HexFromName(coloName));
            return true;
        }

        response = (ColorType.NotDefined, null);
        return false;
    }
}