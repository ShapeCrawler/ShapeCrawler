using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Extensions;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Services;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    internal static class FontDataParser
    {
        public static void GetFontDataFromPlaceholder(ref FontData phFontData, SCParagraph paragraph)
        {
            Shape fontParentShape = (Shape)paragraph.ParentTextBox.TextFrameContainer;
            int paragraphLvl = paragraph.Level;
            if (fontParentShape.Placeholder == null)
            {
                return;
            }

            Placeholder placeholder = (Placeholder)fontParentShape.Placeholder;
            IFontDataReader phReferencedShape = (IFontDataReader)placeholder.ReferencedShape;
            phReferencedShape?.FillFontData(paragraphLvl, ref phFontData);
        }

        /// <summary>
        ///     Gets font data.
        /// </summary>
        /// <param name="compositeElement">Instance of <see cref="P.DefaultTextStyle" /> or <see cref="A.ListStyle" /> class.</param>
        public static Dictionary<int, FontData>
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
                if (aDefRPr?.FontSize != null)
                {
                    fontSize = aDefRPr.FontSize.Value;
                }

                bool? isBold = null;
                if (aDefRPr?.Bold != null)
                {
                    isBold = aDefRPr.Bold.Value;
                }

                bool? isItalic = null;
                if (aDefRPr?.Italic != null)
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
                    A.GradientStop aGradientStop = aDefRPr?.GetFirstChild<A.GradientFill>()?.GradientStopList
                        .GetFirstChild<A.GradientStop>();
                    aRgbColorModelHex = aGradientStop?.RgbColorModelHex;
                    aSchemeColor = aGradientStop?.SchemeColor;
                    aSystemColor = aGradientStop?.SystemColor;
                    aPresetColor = aGradientStop?.PresetColor;
                }

#if NETSTANDARD2_0
                var paragraphLvl = int.Parse(textPr.LocalName[3].ToString(System.Globalization.CultureInfo.CurrentCulture), System.Globalization.CultureInfo.CurrentCulture);
#else
                // fourth character of LocalName contains level number, example: "lvl1pPr -> 1, lvl2pPr -> 2, etc."
                ReadOnlySpan<char> localNameAsSpan = textPr.LocalName.AsSpan();
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
}