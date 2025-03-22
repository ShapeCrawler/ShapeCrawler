using System.Linq;
using ShapeCrawler.Drawing;
using ShapeCrawler.Slides;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal sealed class AutoShape : CopyableShape
{
    private readonly P.Shape pShape;
    private readonly ShapeGeometry shapeGeometry;

    internal AutoShape(P.Shape pShape, TextBox textBox)
        : this(pShape)
    {
        this.IsTextHolder = true;
        this.TextBox = textBox;
    }

    internal AutoShape(P.Shape pShape)
        : base(pShape)
    {
        this.pShape = pShape;
        var shapeProperties = pShape.Descendants<P.ShapeProperties>().First();
        this.Outline = new SlideShapeOutline(shapeProperties);
        this.Fill = new ShapeFill(shapeProperties);
        this.shapeGeometry = new ShapeGeometry(shapeProperties);
    }

    public override bool HasOutline => true;

    public override IShapeOutline Outline { get; }

    public override bool HasFill => true;

    public override IShapeFill Fill { get; }

    public override ShapeType ShapeType => ShapeType.AutoShape;

    public override bool Removeable => true;

    public override Geometry GeometryType
    {
        get => this.shapeGeometry.GeometryType;
        set => this.shapeGeometry.GeometryType = value;
    }

    public override decimal CornerSize
    {
        get => this.shapeGeometry.CornerSize;
        set => this.shapeGeometry.CornerSize = value;
    }

    public override decimal[] Adjustments
    {
        get => this.shapeGeometry.Adjustments;
        set => this.shapeGeometry.Adjustments = value;
    }

    public override void Remove() => this.pShape.Remove();
}