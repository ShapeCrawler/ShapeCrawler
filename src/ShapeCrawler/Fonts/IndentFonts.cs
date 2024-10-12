using System;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Fonts;

internal readonly record struct IndentFonts
{
    private readonly OpenXmlCompositeElement sdkOpenXmlCompositeElement;

    internal IndentFonts(OpenXmlCompositeElement sdkOpenXmlCompositeElement)
    {
        this.sdkOpenXmlCompositeElement = sdkOpenXmlCompositeElement;
    }

    #region APIs

    internal IndentFont? FontOrNull(int indentLevelFor)
    {
        // Get <a:lvlXpPr> elements, eg. <a:lvl1pPr>, <a:lvl2pPr>
        var lvlParagraphPropertyList = this.sdkOpenXmlCompositeElement.Elements()
            .Where(e => e.LocalName.StartsWith("lvl", StringComparison.Ordinal));

        foreach (var textPr in lvlParagraphPropertyList)
        {
            var aDefRPr = textPr.GetFirstChild<A.DefaultRunProperties>();

            int? fontSize = null;
            if (aDefRPr?.FontSize is not null)
            {
                fontSize = aDefRPr.FontSize.Value;
            }

            bool? isBold = null;
            if (aDefRPr?.Bold is not null)
            {
                isBold = aDefRPr.Bold.Value;
            }

            bool? isItalic = null;
            if (aDefRPr?.Italic is not null)
            {
                isItalic = aDefRPr.Italic;
            }

            var aLatinFont = aDefRPr?.GetFirstChild<A.LatinFont>();

            A.RgbColorModelHex? aRgbColorModelHex;
            A.SchemeColor? aSchemeColor;
            A.SystemColor? aSystemColor;
            A.PresetColor? aPresetColor;

            // Try get color from <a:solidFill>
            var aSolidFill = aDefRPr?.SdkASolidFill();
            if (aSolidFill != null)
            {
                aRgbColorModelHex = aSolidFill.RgbColorModelHex!;
                aSchemeColor = aSolidFill.SchemeColor;
                aSystemColor = aSolidFill.SystemColor;
                aPresetColor = aSolidFill.PresetColor;
            }
            else
            {
                var aGradientStop = aDefRPr?.GetFirstChild<A.GradientFill>()?.GradientStopList!
                    .GetFirstChild<A.GradientStop>();
                aRgbColorModelHex = aGradientStop?.RgbColorModelHex;
                aSchemeColor = aGradientStop?.SchemeColor;
                aSystemColor = aGradientStop?.SystemColor;
                aPresetColor = aGradientStop?.PresetColor;
            }

#if NETSTANDARD2_0
            var paragraphLvl =
                int.Parse(
                    textPr.LocalName[3].ToString(System.Globalization.CultureInfo.CurrentCulture),
                    System.Globalization.CultureInfo.CurrentCulture);
#else
            var localName = textPr.LocalName.AsSpan();
            var level = localName.Slice(3, 1); // the fourth character contains level number, eg. "lvl1pPr -> 1, lvl2pPr -> 2, etc."
            var paragraphLvl = int.Parse(level, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.CurrentCulture);
#endif
            if (paragraphLvl == indentLevelFor)
            {
                var indentFont = new IndentFont
                {
                    Size = fontSize,
                    ALatinFont = aLatinFont,
                    IsBold = isBold,
                    IsItalic = isItalic,
                    ARgbColorModelHex = aRgbColorModelHex,
                    ASchemeColor = aSchemeColor,
                    ASystemColor = aSystemColor,
                    APresetColor = aPresetColor
                };

                return indentFont;
            }
        }

        if (indentLevelFor == 1 && this.sdkOpenXmlCompositeElement.Parent is P.TextBody pTextBody)
        {
            var endParaRunPrFs = pTextBody.GetFirstChild<A.Paragraph>() !
                .GetFirstChild<A.EndParagraphRunProperties>()?.FontSize;
            if (endParaRunPrFs is not null)
            {
                var indentFont = new IndentFont
                {
                    Size = endParaRunPrFs
                };

                return indentFont;
            }
        }

        return null;
    }

    internal ColorType? ColorType(int indentLevel)
    {
        var indentFont = this.FontOrNull(indentLevel);
        if (indentFont is null)
        {
            return null;
        }

        if (indentFont.Value.ARgbColorModelHex != null)
        {
            return ShapeCrawler.ColorType.RGB;
        }

        if (indentFont.Value.ASchemeColor != null)
        {
            return ShapeCrawler.ColorType.Theme;
        }

        if (indentFont.Value.ASystemColor != null)
        {
            return ShapeCrawler.ColorType.Standard;
        }

        if (indentFont.Value.APresetColor != null)
        {
            return ShapeCrawler.ColorType.Preset;
        }

        return null;
    }

    internal bool? BoldFlagOrNull(int indentLevel)
    {
        var indentFont = this.FontOrNull(indentLevel);

        return indentFont?.IsBold;
    }

    internal A.LatinFont? ALatinFontOrNull(int indentLevel)
    {
        var indentFont = this.FontOrNull(indentLevel);

        return indentFont?.ALatinFont;
    }

    #endregion APIs
}