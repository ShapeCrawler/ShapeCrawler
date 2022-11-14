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
}