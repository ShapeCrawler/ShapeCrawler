using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Placeholders
{
    internal class FontData
    {
        public FontData(
            Int32Value fontSize, 
            LatinFont aLatinFont, 
            BooleanValue isBold, 
            BooleanValue isItalic,
            A.SchemeColor aSchemeColor) : this(fontSize)
        {
            FontSize = fontSize;
            ALatinFont = aLatinFont;
            IsBold = isBold;
            IsItalic = isItalic;
            ASchemeColor = aSchemeColor;
        }

        public FontData(Int32Value fontSize)
        {
            FontSize = fontSize;
        }

        public Int32Value FontSize { get; }
        public LatinFont ALatinFont { get; }
        public BooleanValue IsBold { get; set; }
        public BooleanValue IsItalic { get; set; }
        public A.SchemeColor ASchemeColor { get; }
    }
}