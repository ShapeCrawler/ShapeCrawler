using System;
using DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Shapes;

internal static class LineTypeMapping
{
    internal static (string Preset, bool Head, bool Tail) ToOpenXml(LineType type) => type switch
    {
        LineType.Line => ("line", false, false),
        LineType.Arrow => ("line", false, true),
        LineType.DoubleArrow => ("line", true, true),
        LineType.ElbowConnector => ("bentConnector2", false, false),
        LineType.ElbowArrowConnector => ("bentConnector2", false, true),
        LineType.ElbowDoubleArrowConnector => ("bentConnector2", true, true),
        LineType.CurvedConnector => ("curvedConnector2", false, false),
        LineType.CurvedArrowConnector => ("curvedConnector2", false, true),
        LineType.CurvedDoubleArrowConnector => ("curvedConnector2", true, true),
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, "The line type is not a connector preset."),
    };

    internal static LineType FromOpenXml(string? preset, LineEndValues? head, LineEndValues? tail)
    {
        var baseType = preset switch
        {
            "bentConnector2" or "bentConnector3" or "bentConnector4" or "bentConnector5" => LineType.ElbowConnector,
            "curvedConnector2" or "curvedConnector3" or "curvedConnector4" or "curvedConnector5" => LineType.CurvedConnector,
            _ => LineType.Line,
        };

        var hasHead = head is not null && head != LineEndValues.None;
        var hasTail = tail is not null && tail != LineEndValues.None;
        return (baseType, hasHead, hasTail) switch
        {
            (LineType.Line, true, true) => LineType.DoubleArrow,
            (LineType.Line, false, true) => LineType.Arrow,
            (LineType.ElbowConnector, false, true) => LineType.ElbowArrowConnector,
            (LineType.ElbowConnector, true, true) => LineType.ElbowDoubleArrowConnector,
            (LineType.CurvedConnector, false, true) => LineType.CurvedArrowConnector,
            (LineType.CurvedConnector, true, true) => LineType.CurvedDoubleArrowConnector,
            _ => baseType,
        };
    }

}
