using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Slides;
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
            return rotation / 60000d;
        }
    }

    public bool HasOutline => true;

    public bool HasFill => true;
    
    public IShape Shape(string groupedShapeName) => this.GroupedShapes.Shape(groupedShapeName);

    public T Shape<T>(string groupedShapeName) =>
        (T)this.GroupedShapes.First(groupedShape => groupedShape is T && groupedShape.Name == groupedShapeName);
}