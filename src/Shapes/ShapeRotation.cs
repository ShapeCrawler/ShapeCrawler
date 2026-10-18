using System;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal sealed class ShapeRotation(OpenXmlElement element)
{
    internal void Set(double degrees)
    {
        if (!double.IsFinite(degrees))
        {
            throw new ArgumentOutOfRangeException(nameof(degrees), "Rotation must be finite.");
        }

        var normalized = degrees % 360d;
        if (normalized < 0)
        {
            normalized += 360d;
        }

        var units = (int)Math.Round(normalized * 60_000d, MidpointRounding.AwayFromZero) % 21_600_000;
        if (element is P.GraphicFrame frame)
        {
            var transform = frame.Transform ?? throw new SCException("The graphic frame has no transform.");
            transform.Rotation = units;
            return;
        }

        if (element is P.GroupShape group)
        {
            var properties = group.GroupShapeProperties ?? throw new SCException("The group has no shape properties.");
            properties.TransformGroup ??= new A.TransformGroup();
            properties.TransformGroup.Rotation = units;
            return;
        }

        var shapeProperties = element.GetFirstChild<P.ShapeProperties>()
                              ?? throw new SCException("The shape has no shape properties.");
        var shapeTransform = shapeProperties.Transform2D;
        if (shapeTransform is null)
        {
            shapeTransform = new A.Transform2D();
            if (element is P.Shape shape && shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                    .GetFirstChild<P.PlaceholderShape>() is not null)
            {
                // Materialize an inherited transform locally; never edit the layout/master.
                var inherited = new ReferencedPShape(element).ATransform2D();
                if (inherited is not null)
                {
                    shapeTransform = (A.Transform2D)inherited.CloneNode(true);
                }
            }

            shapeProperties.AddChild(shapeTransform, true);
        }

        shapeTransform.Rotation = units;
    }
}
