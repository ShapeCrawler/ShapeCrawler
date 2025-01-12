using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Texts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.ShapeCollection;

internal sealed class AutoShape : CopyableShape
{
    private readonly P.Shape pShape;

    private readonly ShapeGeometry shapeGeometry;

    internal AutoShape(
        OpenXmlPart sdkTypedOpenXmlPart,
        P.Shape pShape,
        TextBox textBox)
        : this(sdkTypedOpenXmlPart, pShape)
    {
        this.IsTextHolder = true;
        this.TextBox = textBox;
    }

    internal AutoShape(
        OpenXmlPart sdkTypedOpenXmlPart,
        P.Shape pShape)
        : base(sdkTypedOpenXmlPart, pShape)
    {
        this.pShape = pShape;
        this.Outline = new SlideShapeOutline(this.SdkTypedOpenXmlPart, pShape.Descendants<P.ShapeProperties>().First());
        this.Fill = new ShapeFill(this.SdkTypedOpenXmlPart, pShape.Descendants<P.ShapeProperties>().First());
        this.shapeGeometry = new ShapeGeometry(pShape.Descendants<P.ShapeProperties>().First());
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

    public override void Remove() => this.pShape.Remove();
}