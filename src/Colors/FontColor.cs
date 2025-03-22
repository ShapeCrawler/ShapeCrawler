using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Fonts;
using ShapeCrawler.Paragraphs;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Colors;

internal sealed class FontColor(A.Text aText) : IFontColor
{
    public ColorType Type
    {
        get
        {
            var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
            var aSolidFill = aText.Parent!.GetFirstChild<A.RunProperties>()?.SdkASolidFill();
            if (aSolidFill != null)
            {
                var pSlideMaster = openXmlPart switch
                {
                    SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster,
                    SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideMasterPart!.SlideMaster,
                    _ => ((SlideMasterPart)openXmlPart).SlideMaster
                };
                var typeAndColor = HexParser.FromSolidFill(aSolidFill, pSlideMaster);
                return typeAndColor.Item1;
            }

            var aParagraph = aText.Ancestors<A.Paragraph>().First();
            var indentLevel = new SCAParagraph(aParagraph).GetIndentLevel();
            var pTextBody = aParagraph.Ancestors<P.TextBody>().First();
            var aListStyle = pTextBody.GetFirstChild<A.ListStyle>() !;
            var textBodyStyleFont = new IndentFonts(aListStyle).FontOrNull(indentLevel);
            if (textBodyStyleFont.HasValue && this.TryFromIndentFont(textBodyStyleFont, out var textBodyColor))
            {
                return textBodyColor.colorType;
            }

            // From Shape
            var shapeColor = new ShapeColor(openXmlPart, aText);
            var type = shapeColor.TypeOrNull();
            if (type.HasValue)
            {
                return (ColorType)type;
            }

            // From Referenced Shape
            if (openXmlPart is not SlideMasterPart)
            {
                var refShapeColorType = new ReferencedIndentLevel(aText).ColorTypeOrNull();
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
            var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
            var pSlideMaster = openXmlPart switch
            {
                SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster,
                SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideMasterPart!.SlideMaster,
                _ => ((SlideMasterPart)openXmlPart).SlideMaster
            };
            var aSolidFill = aText.Parent!.GetFirstChild<A.RunProperties>()?.SdkASolidFill();
            if (aSolidFill != null)
            {
                var typeAndColor = HexParser.FromSolidFill(aSolidFill, pSlideMaster);
                return typeAndColor.Item2!;
            }

            // From TextBody
            var aParagraph = aText.Ancestors<A.Paragraph>().First();
            var indentLevel = new SCAParagraph(aParagraph).GetIndentLevel();
            var pTextBody = aParagraph.Ancestors<P.TextBody>().First();
            var textBodyStyleFont = new IndentFonts(pTextBody.GetFirstChild<A.ListStyle>()
                !).FontOrNull(indentLevel);
            if (textBodyStyleFont.HasValue && this.TryFromIndentFont(textBodyStyleFont, out var textBodyColor))
            {
                return textBodyColor.colorHex!;
            }

            // From Shape
            var pShape = aText.Ancestors<P.Shape>().First();
            var sdkSlidePShapeWrap = new ShapeColor(new PresentationColor(openXmlPart), pShape);
            string? shapeFontColorHex = sdkSlidePShapeWrap.HexOrNull();
            if (shapeFontColorHex != null)
            {
                return shapeFontColorHex;
            }

            // From Referenced Shape
            if (openXmlPart is not SlideMasterPart)
            {
                var refShapeFontColorHex = new ReferencedIndentLevel(aText).ColorHexOrNull();
                if (refShapeFontColorHex != null)
                {
                    return refShapeFontColorHex;
                }
            }

            // From Common Placeholder
            var pSlideMasterWrap =
                new SCPSlideMaster(pSlideMaster);
            var masterIndentFont = pSlideMasterWrap.BodyStyleFontOrNull(indentLevel);
            if (this.TryFromIndentFont(masterIndentFont, out var masterColor))
            {
                return masterColor.colorHex!;
            }

            // Presentation level
            var presColor = new PresentationColor(openXmlPart);
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
        var aTextContainer = aText.Parent!;
        var aRunProperties = aTextContainer.GetFirstChild<A.RunProperties>() ?? aTextContainer.AddRunProperties();

        var aSolidFill = aRunProperties.SdkASolidFill();
        aSolidFill?.Remove();
        hex = hex.StartsWith("#", System.StringComparison.Ordinal) ? hex[1..] : hex; // to skip '#'
        var rgbColorModelHex = new A.RgbColorModelHex { Val = hex };
        aSolidFill = new A.SolidFill();
        aSolidFill.Append(rgbColorModelHex);
        aRunProperties.InsertAt(aSolidFill, 0);
    }

    private bool TryFromIndentFont(IndentFont? indentFont, out (ColorType colorType, string? colorHex) response)
    {
        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
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
            var sdkSlidePartWrap = new PresentationColor(openXmlPart);
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