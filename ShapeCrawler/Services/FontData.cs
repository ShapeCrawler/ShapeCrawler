using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Services
{
    internal class FontData
    {
        internal A.SystemColor? ASystemColor { get; set; }
        
        internal Int32Value? FontSize { get; set; }

        internal A.LatinFont? ALatinFont { get; set; }

        internal BooleanValue? IsBold { get; set; }

        internal BooleanValue? IsItalic { get; set; }

        internal A.RgbColorModelHex? ARgbColorModelHex { get; set; }

        internal A.SchemeColor? ASchemeColor { get; set; }

        internal A.PresetColor? APresetColor { get; set; }

        internal bool IsFilled()
        {
            return this.FontSize != null
                   && this.ALatinFont != null
                   && this.IsBold != null
                   && this.IsItalic != null
                   && this.ASchemeColor != null;
        }

        internal void Fill(FontData fontData)
        {
            if (fontData.FontSize == null && this.FontSize != null)
            {
                fontData.FontSize = this.FontSize;
            }

            if (fontData.ALatinFont == null && this.ALatinFont != null)
            {
                fontData.ALatinFont = this.ALatinFont;
            }

            if (fontData.IsBold == null && this.IsBold != null)
            {
                fontData.IsBold = this.IsBold;
            }

            if (fontData.IsItalic == null && this.IsItalic != null)
            {
                fontData.IsItalic = this.IsItalic;
            }

            if (fontData.ASchemeColor == null && this.ASchemeColor != null)
            {
                fontData.ASchemeColor = this.ASchemeColor;
            }

            if (fontData.ARgbColorModelHex == null && this.ARgbColorModelHex != null)
            {
                fontData.ARgbColorModelHex = this.ARgbColorModelHex;
            }

            if (fontData.ASystemColor == null && this.ASystemColor != null)
            {
                fontData.ASystemColor = this.ASystemColor;
            }

            if (fontData.APresetColor == null && this.APresetColor != null)
            {
                fontData.APresetColor = this.APresetColor;
            }
        }
    }
}