using System.Collections.Generic;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Placeholders
{
    /// <summary>
    ///     Represents placeholder font data.
    /// </summary>
    internal class PlaceholderFontData : PlaceholderData
    {
        #region Properties

        internal Dictionary<int, FontData> LvlToFontData;

        #endregion Properties
    }

    internal class FontData
    {
        public FontData(Int32Value fontSize, A.LatinFont aLatinFont) : this(fontSize)
        {
            FontSize = fontSize;
            ALatinFont = aLatinFont;
        }

        public FontData(Int32Value fontSize)
        {
            FontSize = fontSize;
        }

        public Int32Value FontSize { get; }
        public A.LatinFont ALatinFont { get; }
    }
}