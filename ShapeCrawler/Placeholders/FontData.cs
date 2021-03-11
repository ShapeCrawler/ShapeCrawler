using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Placeholders
{
    internal class FontData
    {
        public FontData(Int32Value fontSize, LatinFont aLatinFont) : this(fontSize)
        {
            FontSize = fontSize;
            ALatinFont = aLatinFont;
        }

        public FontData(Int32Value fontSize)
        {
            FontSize = fontSize;
        }

        public Int32Value FontSize { get; }
        public LatinFont ALatinFont { get; }
    }
}