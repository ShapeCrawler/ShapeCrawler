using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Fonts;

internal record struct IndentFont
{
    internal A.SystemColor? ASystemColor { get; init; }

    internal int? Size { get; init; }

    internal A.LatinFont? ALatinFont { get; init; }

    internal bool? IsBold { get; set; }

    internal BooleanValue? IsItalic { get; init; }

    internal A.RgbColorModelHex? ARgbColorModelHex { get; init; }

    internal A.SchemeColor? ASchemeColor { get; init; }

    internal A.PresetColor? APresetColor { get; init; }
}