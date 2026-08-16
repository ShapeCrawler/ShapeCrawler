using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ShapeCrawler.Units;

namespace ShapeCrawler.Slides;

internal static class CustomLineShape
{
    internal static string Xml(LineType type, IReadOnlyList<Point> points, int id)
    {
        if (points is null || points.Count < 2)
        {
            throw new ArgumentException("At least two points are required to create a line.", nameof(points));
        }

        var minX = points.Min(point => point.X);
        var minY = points.Min(point => point.Y);
        var maxX = points.Max(point => point.X);
        var maxY = points.Max(point => point.Y);
        var width = Math.Max(1m, maxX - minX);
        var height = Math.Max(1m, maxY - minY);
        var widthEmu = new Points(width).AsEmus();
        var heightEmu = new Points(height).AsEmus();
        var path = new StringBuilder();
        var first = points[0];
        path.Append($"<a:moveTo><a:pt x=\"{Emu(first.X - minX)}\" y=\"{Emu(first.Y - minY)}\"/></a:moveTo>");
        if (type == LineType.Curve && points.Count >= 3)
        {
            for (var i = 1; i + 1 < points.Count; i += 2)
            {
                var control = points[i];
                var end = points[i + 1];
                path.Append($"<a:quadBezTo><a:pt x=\"{Emu(control.X - minX)}\" y=\"{Emu(control.Y - minY)}\"/><a:pt x=\"{Emu(end.X - minX)}\" y=\"{Emu(end.Y - minY)}\"/></a:quadBezTo>");
            }

            if (points.Count % 2 == 0)
            {
                var last = points[^1];
                path.Append($"<a:lnTo><a:pt x=\"{Emu(last.X - minX)}\" y=\"{Emu(last.Y - minY)}\"/></a:lnTo>");
            }
        }
        else
        {
            foreach (var point in points.Skip(1))
            {
                path.Append($"<a:lnTo><a:pt x=\"{Emu(point.X - minX)}\" y=\"{Emu(point.Y - minY)}\"/></a:lnTo>");
            }
        }

        var name = type switch
        {
            LineType.Curve => "Curve",
            LineType.FreeformShape => "Freeform Shape",
            LineType.Scribble => "Scribble",
            _ => throw new ArgumentOutOfRangeException(nameof(type), type, "Custom line type expected."),
        };

        return $"""
            <p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
              <p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
              <p:spPr>
                <a:xfrm><a:off x="{Emu(minX)}" y="{Emu(minY)}"/><a:ext cx="{widthEmu}" cy="{heightEmu}"/></a:xfrm>
                <a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>
                  <a:pathLst><a:path w="{widthEmu}" h="{heightEmu}" fill="none" stroke="1">{path}</a:path></a:pathLst>
                </a:custGeom>
                <a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
              </p:spPr>
            </p:sp>
            """;
    }

    private static long Emu(decimal value) => new Points(value).AsEmus();
}