using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Services;

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
        return this.FontSize is not null
               && this.ALatinFont is not null
               && this.IsBold is not null
               && this.IsItalic is not null
               && this.ASchemeColor is not null;
    }

    internal void Fill(FontData fontData)
    {
        if (fontData.FontSize is null && this.FontSize is not null )
        {
            fontData.FontSize = this.FontSize;
        }

        if (fontData.ALatinFont == null && this.ALatinFont != null)
        {
            fontData.ALatinFont = this.ALatinFont;
        }

        if (fontData.IsBold is null && this.IsBold is not null )
        {
            fontData.IsBold = this.IsBold;
        }

        if (fontData.IsItalic is null && this.IsItalic is not null )
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