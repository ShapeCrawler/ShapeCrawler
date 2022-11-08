using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Services;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Factories;

internal static class FontDataParser
{
    internal static void GetFontDataFromPlaceholder(ref FontData phFontData, SCParagraph paragraph)
    {
        var shape = (Shape)paragraph.ParentTextFrame.TextFrameContainer;
        if (shape.Placeholder == null)
        {
            return;
        }

        var placeholder = (Placeholder)shape.Placeholder;
        var autoShape = (SlideAutoShape)placeholder.ReferencedShape;
        autoShape?.FillFontData(paragraph.Level, ref phFontData);
    }

    internal static Dictionary<int, FontData>
        FromCompositeElement(
            OpenXmlCompositeElement compositeElement)
    {
        // Get <a:lvlXpPr> elements, eg. <a:lvl1pPr>, <a:lvl2pPr>
        IEnumerable<OpenXmlElement> lvlParagraphPropertyList = compositeElement.Elements()
            .Where(e => e.LocalName.StartsWith("lvl", StringComparison.Ordinal));

        var lvlToFontData = new Dictionary<int, FontData>();
        foreach (OpenXmlElement textPr in lvlParagraphPropertyList)
        {
            A.DefaultRunProperties aDefRPr = textPr.GetFirstChild<A.DefaultRunProperties>();

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

            A.LatinFont aLatinFont = aDefRPr?.GetFirstChild<A.LatinFont>();

            A.RgbColorModelHex aRgbColorModelHex;
            A.SchemeColor aSchemeColor;
            A.SystemColor aSystemColor;
            A.PresetColor aPresetColor;

            // Try get color from <a:solidFill>
            A.SolidFill aSolidFill = aDefRPr?.GetASolidFill();
            if (aSolidFill != null)
            {
                aRgbColorModelHex = aSolidFill.RgbColorModelHex;
                aSchemeColor = aSolidFill.SchemeColor;
                aSystemColor = aSolidFill.SystemColor;
                aPresetColor = aSolidFill.PresetColor;
            }
            else
            {
                A.GradientStop aGradientStop = aDefRPr?.GetFirstChild<A.GradientFill>()?.GradientStopList!
                    .GetFirstChild<A.GradientStop>();
                aRgbColorModelHex = aGradientStop?.RgbColorModelHex;
                aSchemeColor = aGradientStop?.SchemeColor;
                aSystemColor = aGradientStop?.SystemColor;
                aPresetColor = aGradientStop?.PresetColor;
            }

#if NETSTANDARD2_0
                var paragraphLvl =
 int.Parse(textPr.LocalName[3].ToString(System.Globalization.CultureInfo.CurrentCulture), System.Globalization.CultureInfo.CurrentCulture);
#else
            // fourth character of LocalName contains level number, example: "lvl1pPr -> 1, lvl2pPr -> 2, etc."
            var localNameAsSpan = textPr.LocalName.AsSpan();
            int paragraphLvl = int.Parse(localNameAsSpan.Slice(3, 1));
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
            lvlToFontData.Add(paragraphLvl, fontData);
        }

        return lvlToFontData;
    }
}