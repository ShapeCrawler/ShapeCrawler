using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Groups;
using ShapeCrawler.Positions;

namespace ShapeCrawler.Shapes;
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