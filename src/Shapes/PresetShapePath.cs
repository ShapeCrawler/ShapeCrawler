using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using SkiaSharp;

namespace ShapeCrawler.Shapes;

internal sealed class PresetShapePath(PresetGeometry? preset, float width, float height)
{
    internal SKPath Create(Geometry geometry)
    {
        var shortest = Math.Min(width, height);
        var halfWidth = width / 2;
        var halfHeight = height / 2;
        var inset = shortest * this.Adjustment("adj", 25000, 0, 100000 * width / shortest) / 100000;
        return geometry switch
        {
            Geometry.Triangle => Polygon((0, height),
                (width * this.Adjustment("adj", 50000, 0, 100000) / 100000, 0), (width, height)),
            Geometry.RightTriangle => Polygon((0, height), (0, 0), (width, height)),
            Geometry.Diamond => Polygon((0, halfHeight), (halfWidth, 0), (width, halfHeight), (halfWidth, height)),
            Geometry.Parallelogram => Polygon((0, height), (inset, 0), (width, 0), (width - inset, height)),
            Geometry.Trapezoid => this.Trapezoid(shortest),
            Geometry.Pentagon => this.Pentagon(),
            Geometry.Hexagon => this.Hexagon(shortest),
            Geometry.Octagon => this.Octagon(shortest),
            Geometry.Plus => this.Plus(shortest),
            Geometry.Chevron => this.Chevron(shortest, false),
            Geometry.HomePlate => this.Chevron(shortest, true),
            Geometry.RightArrow => this.Arrow(false, false, false),
            Geometry.LeftArrow => this.Arrow(false, true, false),
            Geometry.DownArrow => this.Arrow(true, false, false),
            Geometry.UpArrow => this.Arrow(true, true, false),
            Geometry.LeftRightArrow => this.Arrow(false, false, true),
            Geometry.UpDownArrow => this.Arrow(true, false, true),
            _ => throw new SCException("Unsupported shape geometry type.")
        };
    }

    private static SKPath Polygon(params (float X, float Y)[] points)
    {
        using var path = new SKPathBuilder();
        path.MoveTo(points[0].X, points[0].Y);
        foreach (var (x, y) in points.Skip(1))
        {
            path.LineTo(x, y);
        }

        path.Close();
        return path.Detach();
    }

    private float Adjustment(string name, float fallback, float minimum, float maximum)
    {
        var formula = preset?.AdjustValueList?.Elements<ShapeGuide>()
            .FirstOrDefault(guide => guide.Name?.Value == name)?.Formula?.Value;
        var value = formula?.StartsWith("val ", StringComparison.Ordinal) == true &&
            float.TryParse(formula.AsSpan(4), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
            ? parsed : fallback;
        return Math.Clamp(value, minimum, maximum);
    }

    private SKPath Trapezoid(float shortest)
    {
        var inset = shortest * this.Adjustment("adj", 25000, 0, 50000 * width / shortest) / 100000;
        return Polygon((0, height), (inset, 0), (width - inset, 0), (width, height));
    }

    private SKPath Pentagon()
    {
        var horizontal = width / 2 * this.Adjustment("hf", 105146, float.MinValue, float.MaxValue) / 100000;
        var vertical = height / 2 * this.Adjustment("vf", 110557, float.MinValue, float.MaxValue) / 100000;
        var dx1 = horizontal * MathF.Cos(MathF.PI / 10);
        var dx2 = horizontal * MathF.Cos(3 * MathF.PI / 10);
        var y1 = vertical * (1 - MathF.Sin(MathF.PI / 10));
        var y2 = vertical * (1 + MathF.Sin(3 * MathF.PI / 10));
        return Polygon((width / 2 - dx1, y1), (width / 2, 0), (width / 2 + dx1, y1),
            (width / 2 + dx2, y2), (width / 2 - dx2, y2));
    }

    private SKPath Hexagon(float shortest)
    {
        var inset = shortest * this.Adjustment("adj", 25000, 0, 50000 * width / shortest) / 100000;
        var dy = height / 2 * this.Adjustment("vf", 115470, float.MinValue, float.MaxValue) / 100000 * MathF.Sin(MathF.PI / 3);
        return Polygon((0, height / 2), (inset, height / 2 - dy), (width - inset, height / 2 - dy),
            (width, height / 2), (width - inset, height / 2 + dy), (inset, height / 2 + dy));
    }

    private SKPath Octagon(float shortest)
    {
        var inset = shortest * this.Adjustment("adj", 29289, 0, 50000) / 100000;
        return Polygon((0, inset), (inset, 0), (width - inset, 0), (width, inset),
            (width, height - inset), (width - inset, height), (inset, height), (0, height - inset));
    }

    private SKPath Plus(float shortest)
    {
        var inset = shortest * this.Adjustment("adj", 25000, 0, 50000) / 100000;
        return Polygon((0, inset), (inset, inset), (inset, 0), (width - inset, 0),
            (width - inset, inset), (width, inset), (width, height - inset),
            (width - inset, height - inset), (width - inset, height), (inset, height),
            (inset, height - inset), (0, height - inset));
    }

    private SKPath Chevron(float shortest, bool homePlate)
    {
        var inset = shortest * this.Adjustment("adj", 50000, 0, 100000 * width / shortest) / 100000;
        return homePlate
            ? Polygon((0, 0), (width - inset, 0), (width, height / 2), (width - inset, height), (0, height))
            : Polygon((0, 0), (width - inset, 0), (width, height / 2),
                (width - inset, height), (0, height), (inset, height / 2));
    }

    private SKPath Arrow(bool vertical, bool reverse, bool doubleHeaded)
    {
        var length = vertical ? height : width;
        var breadth = vertical ? width : height;
        var shortest = Math.Min(width, height);
        var head = shortest * this.Adjustment("adj2", 50000, 0,
            (doubleHeaded ? 50000 : 100000) * length / shortest) / 100000;
        var halfShaft = breadth * this.Adjustment("adj1", 50000, 0, 100000) / 200000;
        var top = breadth / 2 - halfShaft;
        var bottom = breadth / 2 + halfShaft;
        using var horizontal = doubleHeaded
            ? Polygon((0, breadth / 2), (head, 0), (head, top), (length - head, top),
                (length - head, 0), (length, breadth / 2), (length - head, breadth),
                (length - head, bottom), (head, bottom), (head, breadth))
            : Polygon((0, top), (length - head, top), (length - head, 0), (length, breadth / 2),
                (length - head, breadth), (length - head, bottom), (0, bottom));
        var path = new SKPath(horizontal);
        var transform = vertical
            ? new SKMatrix { ScaleX = 0, SkewX = 1, SkewY = reverse ? -1 : 1,
                ScaleY = 0, TransY = reverse ? height : 0, Persp2 = 1 }
            : SKMatrix.CreateScale(reverse ? -1 : 1, 1);
        if (!vertical && reverse)
        {
            transform.TransX = width;
        }

        path.Transform(transform);
        return path;
    }
}
