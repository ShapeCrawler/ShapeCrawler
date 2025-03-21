using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.GroupShapes;

internal sealed class GroupShape : Shape, IGroupShape
{
    private readonly P.GroupShape pGroupShape;

    internal GroupShape(P.GroupShape pGroupShape)
        : base(pGroupShape)
    {
        this.pGroupShape = pGroupShape;
        this.Shapes = new GroupedShapeCollection(pGroupShape.Elements<OpenXmlCompositeElement>());
        this.Outline = new SlideShapeOutline(pGroupShape.Descendants<P.ShapeProperties>().First());
        this.Fill = new ShapeFill(pGroupShape.Descendants<P.ShapeProperties>().First());
    }

    public IShapeCollection Shapes { get; }
    
    public override Geometry GeometryType => Geometry.Rectangle;
    
    public override ShapeType ShapeType => ShapeType.Group;
    
    public override bool HasOutline => true;
    
    public override IShapeOutline Outline { get; }
    
    public override bool HasFill => true;
    
    public override IShapeFill Fill { get; }
    
    public override bool Removeable => true;
    
    public override double Rotation
    {
        get
        {
            var aTransformGroup = this.pGroupShape.GroupShapeProperties!.TransformGroup!;
            var rotation = aTransformGroup.Rotation?.Value ?? 0;
            return rotation / 60000d;
        }
    }
    
    public override void Remove() => this.pGroupShape.Remove();
}