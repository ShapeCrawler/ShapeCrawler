using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Placeholders
{
    internal class FontData // TODO: can be structure?
    {
        // TODO: remove unnecessary public setter for property below
        public Int32Value FontSize { get; set; }

        public A.LatinFont ALatinFont { get; set; }

        public BooleanValue IsBold { get; set; }

        public BooleanValue IsItalic { get; set; }

        public A.RgbColorModelHex ARgbColorModelHex { get; set; }

        public A.SchemeColor ASchemeColor { get; set; }

        public A.SystemColor ASystemColor { get; set; }

        public A.PresetColor APresetColor { get; set; }

        public bool IsFilled()
        {
            return this.FontSize != null
                   && this.ALatinFont != null
                   && this.IsBold != null
                   && this.IsItalic != null
                   && this.ASchemeColor != null;
        }

        public void Fill(FontData fontData)
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