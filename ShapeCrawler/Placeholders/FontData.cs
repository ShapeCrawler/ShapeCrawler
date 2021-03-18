using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Placeholders
{
    internal class FontData
    {
        public FontData(Int32Value fontSize, LatinFont aLatinFont, bool? isBold) : this(fontSize)
        {
            FontSize = fontSize;
            ALatinFont = aLatinFont;
            IsBold = isBold;
        }

        public FontData(Int32Value fontSize)
        {
            FontSize = fontSize;
        }

        public Int32Value FontSize { get; }
        public LatinFont ALatinFont { get; }
        public bool? IsBold { get; }
    }
}