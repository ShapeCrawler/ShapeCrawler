using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Services;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal sealed record PSlideMasterWrap
{
    private readonly P.SlideMaster pSlideMaster;

    internal PSlideMasterWrap(P.SlideMaster pSlideMaster)
    {
        this.pSlideMaster = pSlideMaster;
    }

    internal FontData? BodyStyleParagraphLevelFontOrNull(int paragraphLvlParam)
    {
        // Get <a:lvlXpPr> elements, eg. <a:lvl1pPr>, <a:lvl2pPr>
        var lvlParagraphPropertyList = this.pSlideMaster.TextStyles!.BodyStyle!.Elements()
            .Where(e => e.LocalName.StartsWith("lvl", StringComparison.Ordinal));

        foreach (OpenXmlElement textPr in lvlParagraphPropertyList)
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
            var aSolidFill = aDefRPr?.GetASolidFill();
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
            var level =
                localName.Slice(3,
                    1); // the fourth character contains level number, eg. "lvl1pPr -> 1, lvl2pPr -> 2, etc."
            var paragraphLvl =
                int.Parse(level, System.Globalization.NumberStyles.Number,
                    System.Globalization.CultureInfo.CurrentCulture);
#endif
            var fontData = new FontData
            {
                FontSize = fontSize,
                ALatinFont = aLatinFont,
                IsBold = isBold,
                IsItalic = isItalic,
                ARgbColorModelHex = aRgbColorModelHex,
                ASchemeColor = aSchemeColor,
                ASystemColor = aSystemColor,
                APresetColor = aPresetColor
            };

            if (paragraphLvl == paragraphLvlParam)
            {
                return fontData;
            }
        }

        return null;
    }
}