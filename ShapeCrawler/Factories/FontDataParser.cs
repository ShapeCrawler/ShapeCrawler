using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Placeholders;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    internal static class FontDataParser
    {
        public static void GetFontDataFromPlaceholder(ref FontData phFontData, SCParagraph paragraph)
        {
            Shape fontParentShape = (Shape)paragraph.ParentTextBox.ParentTextBoxContainer;
            int paragraphLvl = paragraph.Level;
            if (fontParentShape.Placeholder == null)
            {
                return;
            }

            Placeholder placeholder = (Placeholder) fontParentShape.Placeholder;
            IFontDataReader phReferencedShape = (IFontDataReader) placeholder.ReferencedShape;
            phReferencedShape?.FillFontData(paragraphLvl, ref phFontData);
        }

        /// <summary>
        ///     Gets font data.
        /// </summary>
        /// <param name="compositeElement">Instance of <see cref="P.DefaultTextStyle" /> or <see cref="A.ListStyle" /> class.</param>
        public static Dictionary<int, FontData>
            FromCompositeElement(
                OpenXmlCompositeElement compositeElement) // TODO: set annotation that about it cannot be NULL
        {
            // Get <a:lvlXpPr> elements, eg. <a:lvl1pPr>, <a:lvl2pPr>
            IEnumerable<OpenXmlElement> lvlParagraphPropertyList = compositeElement.Elements()
                .Where(e => e.LocalName.StartsWith("lvl", StringComparison.Ordinal));

            var lvlToFontData = new Dictionary<int, FontData>();
            foreach (OpenXmlElement textPr in lvlParagraphPropertyList)
            {
                A.DefaultRunProperties aDefRPr = textPr.GetFirstChild<A.DefaultRunProperties>();

                Int32Value fontSize = aDefRPr?.FontSize;
                BooleanValue isBold = aDefRPr?.Bold;
                BooleanValue isItalic = aDefRPr?.Italic;
                A.LatinFont aLatinFont = aDefRPr?.GetFirstChild<A.LatinFont>();
                A.RgbColorModelHex aRgbColorModelHex = aDefRPr?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex;
                A.SchemeColor aSchemeColor = aDefRPr?.GetFirstChild<A.SolidFill>()?.SchemeColor;

#if NETSTANDARD2_0
                var lvl = int.Parse(textPr.LocalName[3].ToString(System.Globalization.CultureInfo.CurrentCulture), System.Globalization.CultureInfo.CurrentCulture);
#else
                // fourth character of LocalName contains level number, example: "lvl1pPr -> 1, lvl2pPr -> 2, etc."
                ReadOnlySpan<char> localNameAsSpan = textPr.LocalName.AsSpan();
                int lvl = int.Parse(localNameAsSpan.Slice(3, 1));
#endif
                lvlToFontData.Add(lvl, new FontData(fontSize, aLatinFont, isBold, isItalic, aRgbColorModelHex, aSchemeColor));
            }

            return lvlToFontData;
        }
    }
}