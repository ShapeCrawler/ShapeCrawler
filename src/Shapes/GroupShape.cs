using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Groups;
using ShapeCrawler.Positions;

namespace ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

internal sealed class GroupShape : Shape
{
    private readonly P.GroupShape pGroupShape;

    internal GroupShape(P.GroupShape pGroupShape): base(new Position(pGroupShape), new ShapeSize(pGroupShape), new ShapeId(pGroupShape), pGroupShape)
    {
        this.pGroupShape = pGroupShape;
        this.Shapes = new GroupedShapeCollection(pGroupShape.Elements<OpenXmlCompositeElement>());
        var pShapeProperties = pGroupShape.Descendants<P.ShapeProperties>().First();
    }

    public IShapeCollection Shapes { get; }

    public override Geometry GeometryType => Geometry.Rectangle;

    public bool HasOutline => true;

    public bool HasFill => true;
    
    public IShape Shape(string groupedShapeName) => this.Shapes.Shape(groupedShapeName);

    public T Shape<T>(string groupedShapeName) =>
        (T)this.Shapes.First(groupedShape => groupedShape is T && groupedShape.Name == groupedShapeName);
}