using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.SlideShape;

internal sealed class SlideGroupShape : Shape, IGroupShape
{
    private readonly P.GroupShape pGroupShape;

    internal SlideGroupShape(TypedOpenXmlPart sdkTypedOpenXmlPart, P.GroupShape pGroupShape)
        : base(sdkTypedOpenXmlPart, pGroupShape)
    {
        this.pGroupShape = pGroupShape;
        this.Shapes = new SlideGroupedShapes(sdkTypedOpenXmlPart, pGroupShape.Elements<OpenXmlCompositeElement>());
        this.Outline = new SlideShapeOutline(sdkTypedOpenXmlPart, pGroupShape.Descendants<P.ShapeProperties>().First());
        this.Fill = new SlideShapeFill(sdkTypedOpenXmlPart, pGroupShape.Descendants<P.ShapeProperties>().First(), false);
    }

    public IReadOnlyShapes Shapes { get; }
    public override Geometry GeometryType => Geometry.Rectangle;
    public override ShapeType ShapeType => ShapeType.Group;
    public override bool HasOutline => true;
    public override IShapeOutline Outline { get; }
    public override bool HasFill => true;
    public override IShapeFill Fill { get; }
    public override bool Removeable => true;
    public override void Remove()=> this.pGroupShape.Remove();
}