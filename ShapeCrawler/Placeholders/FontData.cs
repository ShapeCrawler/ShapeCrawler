using DocumentFormat.OpenXml;

namespace ShapeCrawler.Placeholders
{
    internal class FontData
    {
        public FontData(Int32Value fontSize, DocumentFormat.OpenXml.Drawing.LatinFont aLatinFont) : this(fontSize)
        {
            FontSize = fontSize;
            ALatinFont = aLatinFont;
        }

        public FontData(Int32Value fontSize)
        {
            FontSize = fontSize;
        }

        public Int32Value FontSize { get; }
        public DocumentFormat.OpenXml.Drawing.LatinFont ALatinFont { get; }
    }
}