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

internal sealed class SlideGroupShape : Shape, IGroupShape, IRemoveable
{
    private readonly P.GroupShape pGroupShape;

    internal SlideGroupShape(SlidePart sdkSlidePart, P.GroupShape pGroupShape)
        : base(pGroupShape)
    {
        this.pGroupShape = pGroupShape;
        this.Shapes = new SlideGroupedShapes(sdkSlidePart, pGroupShape.Elements<OpenXmlCompositeElement>());
        this.Outline = new SlideShapeOutline(sdkSlidePart, pGroupShape.Descendants<P.ShapeProperties>().First());
        this.Fill = new SlideShapeFill(sdkSlidePart, pGroupShape.Descendants<P.ShapeProperties>().First(), false);
    }

    public IReadOnlyShapes Shapes { get; }
    public override SCGeometry GeometryType => SCGeometry.Rectangle;
    public override SCShapeType ShapeType => SCShapeType.Group;
    public override bool HasOutline => true;
    public override IShapeOutline Outline { get; }
    public override bool HasFill => true;
    public override IShapeFill Fill { get; }
    void  IRemoveable.Remove() => this.pGroupShape.Remove();
}