﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Placeholders
{
    internal class FontData
    {
        public FontData(Int32Value fontSize, LatinFont aLatinFont, BooleanValue isBold, BooleanValue isItalic) : this(fontSize)
        {
            FontSize = fontSize;
            ALatinFont = aLatinFont;
            IsBold = isBold;
            IsItalic = isItalic;
        }

        public FontData(Int32Value fontSize)
        {
            FontSize = fontSize;
        }

        public Int32Value FontSize { get; }
        public LatinFont ALatinFont { get; }
        public BooleanValue IsBold { get; set; }
        public BooleanValue IsItalic { get; set; }
    }
}