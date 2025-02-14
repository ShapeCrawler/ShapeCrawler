using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Fonts;

internal record struct IndentFont
{
    internal A.SystemColor? ASystemColor { get; set; }

    internal int? Size { get; set; }

    internal A.LatinFont? ALatinFont { get; set; }

    internal bool? IsBold { get; set; }

    internal BooleanValue? IsItalic { get; set; }

    internal A.RgbColorModelHex? ARgbColorModelHex { get; set; }

    internal A.SchemeColor? ASchemeColor { get; set; }

    internal A.PresetColor? APresetColor { get; set; }
}