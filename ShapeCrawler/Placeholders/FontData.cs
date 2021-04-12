using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Placeholders
{
    internal class FontData // TODO: can be structure?
    {
        public FontData(
            Int32Value fontSize,
            A.LatinFont aLatinFont,
            BooleanValue isBold,
            BooleanValue isItalic,
            A.RgbColorModelHex aRgbColorModelHex,
            A.SchemeColor aSchemeColor) : this(fontSize)
        {
            FontSize = fontSize;
            ALatinFont = aLatinFont;
            IsBold = isBold;
            IsItalic = isItalic;
            ASchemeColor = aSchemeColor;
            ARgbColorModelHex = aRgbColorModelHex;
        }

        public FontData()
        {
        }

        public FontData(Int32Value fontSize)
        {
            FontSize = fontSize;
        }

        // TODO: remove unnecessary public setter for property below
        public Int32Value FontSize { get; set; }
        public A.LatinFont ALatinFont { get; set; }
        public BooleanValue IsBold { get; set; }
        public BooleanValue IsItalic { get; set; }
        public A.RgbColorModelHex ARgbColorModelHex { get; set; }
        public A.SchemeColor ASchemeColor { get; set; }

        public bool IsFilled()
        {
            return FontSize != null && ALatinFont != null && IsBold != null && IsItalic != null && ASchemeColor != null;
        }

        public void Fill(FontData fontData)
        {
            if (fontData.FontSize == null && FontSize != null)
            {
                fontData.FontSize = FontSize;
            }

            if (fontData.ALatinFont == null && ALatinFont != null)
            {
                fontData.ALatinFont = ALatinFont;
            }

            if (fontData.IsBold == null && IsBold != null)
            {
                fontData.IsBold = IsBold;
            }

            if (fontData.IsItalic == null && IsItalic != null)
            {
                fontData.IsItalic = IsItalic;
            }

            if (fontData.ASchemeColor == null && ASchemeColor != null)
            {
                fontData.ASchemeColor = ASchemeColor;
            }

            if (fontData.ARgbColorModelHex == null && ARgbColorModelHex != null)
            {
                fontData.ARgbColorModelHex = ARgbColorModelHex;
            }
        }

        public void FillSize(int fontSize)
        {
            if (FontSize == null)
            {
                FontSize = new Int32Value(fontSize);
            }
        }
    }
}