using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Slides;
using ShapeCrawler.Units;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Groups;
using P = DocumentFormat.OpenXml.Presentation;

internal sealed class GroupShape : Shape
{
    private readonly P.GroupShape pGroupShape;

    internal GroupShape(P.GroupShape pGroupShape)
        : base(new Position(pGroupShape), new ShapeSize(pGroupShape), new ShapeId(pGroupShape), pGroupShape)
    {
        this.pGroupShape = pGroupShape;
        this.GroupedShapes = new GroupedShapeCollection(pGroupShape.Elements<OpenXmlCompositeElement>());
    }

    internal GroupShape(P.GroupShape pGroupShape, IShape[] groupingShapes, NewShapeProperties idGenerator, SlidePart slidePart) 
        : base(new Position(pGroupShape), new ShapeSize(pGroupShape), new ShapeId(pGroupShape), pGroupShape)
    {
        var nonVisualGroupShapeProperties = new P.NonVisualGroupShapeProperties();
        var id = idGenerator.Id();
        var name = $"Shape {id}";
        var nonVisualDrawingProperties = new P.NonVisualDrawingProperties
        {
            Id = (uint)id, Name = name
        };
        var nonVisualGroupShapeDrawingProperties = new P.NonVisualGroupShapeDrawingProperties();
        var applicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties();

        nonVisualGroupShapeProperties.Append(nonVisualDrawingProperties);
        nonVisualGroupShapeProperties.Append(nonVisualGroupShapeDrawingProperties);
        nonVisualGroupShapeProperties.Append(applicationNonVisualDrawingProperties);

        var groupShapeProperties = new P.GroupShapeProperties();

        decimal minX = decimal.MaxValue;
        decimal minY = decimal.MaxValue;
        decimal maxX = decimal.MinValue;
        decimal maxY = decimal.MinValue;

        foreach (var groupingShape in groupingShapes)
        {
            minX = Math.Min(minX, groupingShape.X);
            minY = Math.Min(minY, groupingShape.Y);
            maxX = Math.Max(maxX, groupingShape.X + groupingShape.Width);
            maxY = Math.Max(maxY, groupingShape.Y + groupingShape.Height);
        }

        var transformGroup = new A.TransformGroup();
        var offset = new A.Offset { X = (int)minX, Y = (int)minY };
        var extents = new A.Extents { Cx = (int)(maxX - minX), Cy = (int)(maxY - minY) };
        var childOffset = new A.ChildOffset { X = 0, Y = 0 };
        var childExtents = new A.ChildExtents { Cx = extents.Cx, Cy = extents.Cy };

        transformGroup.Append(offset);
        transformGroup.Append(extents);
        transformGroup.Append(childOffset);
        transformGroup.Append(childExtents);

        groupShapeProperties.Append(transformGroup);

        pGroupShape.Append(nonVisualGroupShapeProperties);
        pGroupShape.Append(groupShapeProperties);

        foreach (var openXmlElement in groupingShapes.Select(groupingShape => groupingShape.SDKOpenXmlElement))
        {
            // Remove the shape from its current parent
            if (openXmlElement.Parent is not null)
            {
                openXmlElement.Remove();
            }
            
            pGroupShape.Append(openXmlElement);
        }

        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pGroupShape);

        foreach (var grouping in groupingShapes)
        {
            grouping.Remove();
        }
        
        this.pGroupShape = pGroupShape;
        this.GroupedShapes = new GroupedShapeCollection(pGroupShape.Elements<OpenXmlCompositeElement>());
    }

    public override Geometry GeometryType => Geometry.Rectangle;
    
    public override IShapeCollection GroupedShapes { get; }
    
    public override double Rotation
    {
        get
        {
            var aTransformGroup = this.pGroupShape.GroupShapeProperties!.TransformGroup!;
            var rotation = aTransformGroup.Rotation?.Value ?? 0;
            return rotation / 60_000d;
        }
    }

    public bool HasOutline => true;

    public bool HasFill => true;
    
    public IShape Shape(string groupedShapeName) => this.GroupedShapes.Shape(groupedShapeName);

    public T Shape<T>(string groupedShapeName) =>
        (T)this.GroupedShapes.First(groupedShape => groupedShape is T && groupedShape.Name == groupedShapeName);

    internal override void Render(SKCanvas canvas)
    {
        canvas.Save();
        this.ApplyRotation(canvas);
        this.RenderGroupedShapes(canvas);
        canvas.Restore();
    }

    private void ApplyRotation(SKCanvas canvas)
    {
        const double epsilon = 1e-6;
        if (Math.Abs(this.Rotation) <= epsilon)
        {
            return;
        }

        var (x, y, width, height) = this.AbsoluteBounds();
        var centerX = x + (width / 2);
        var centerY = y + (height / 2);
        canvas.RotateDegrees(
            (float)this.Rotation,
            (float)new Points(centerX).AsPixels(),
            (float)new Points(centerY).AsPixels()
        );
    }

    private void RenderGroupedShapes(SKCanvas canvas)
    {
        foreach (var shape in this.GroupedShapes)
        {
            if (shape.Hidden)
            {
                continue;
            }

            if (shape is Shape internalShape)
            {
                internalShape.Render(canvas);
            }
        }
    }

    private (decimal X, decimal Y, decimal Width, decimal Height) AbsoluteBounds()
    {
        var pGroupShapes = this.pGroupShape.Ancestors<P.GroupShape>().ToArray();
        if (pGroupShapes.Length == 0)
        {
            return (this.X, this.Y, this.Width, this.Height);
        }

        decimal absoluteX = this.X;
        decimal absoluteY = this.Y;
        decimal scaleFactorX = 1.0m;
        decimal scaleFactorY = 1.0m;

        foreach (var childPGroupShape in pGroupShapes)
        {
            var transformGroup = childPGroupShape.GroupShapeProperties!.TransformGroup!;
            var childOffset = transformGroup.ChildOffset!;
            var childExtents = transformGroup.ChildExtents!;
            var offset = transformGroup.Offset!;
            var extents = transformGroup.Extents!;

            decimal currentScaleFactorX = 1.0m;
            if (childExtents.Cx!.Value != 0)
            {
                currentScaleFactorX = (decimal)extents.Cx!.Value / childExtents.Cx!.Value;
            }

            decimal currentScaleFactorY = 1.0m;
            if (childExtents.Cy!.Value != 0)
            {
                currentScaleFactorY = (decimal)extents.Cy!.Value / childExtents.Cy!.Value;
            }

            var childOffsetX = new Emus(childOffset.X!.Value).AsPoints();
            var childOffsetY = new Emus(childOffset.Y!.Value).AsPoints();
            absoluteX = ((absoluteX - childOffsetX) * currentScaleFactorX) + new Emus(offset.X!.Value).AsPoints();
            absoluteY = ((absoluteY - childOffsetY) * currentScaleFactorY) + new Emus(offset.Y!.Value).AsPoints();
            scaleFactorX *= currentScaleFactorX;
            scaleFactorY *= currentScaleFactorY;
        }

        var absoluteWidth = this.Width * scaleFactorX;
        var absoluteHeight = this.Height * scaleFactorY;

        return (absoluteX, absoluteY, absoluteWidth, absoluteHeight);
    }
}