using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using SkiaSharp;
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
    
    internal void Draw(SKCanvas slideCanvas)
    {
        var skColorOutline = SKColor.Parse(this.Outline.HexColor);

        using var paint = new SKPaint
        {
            Color = skColorOutline,
            IsAntialias = true,
            StrokeWidth = (float)UnitConverter.PointToPixel(this.Outline.Weight),
            Style = SKPaintStyle.Stroke
        };

        if (this.GeometryType == Geometry.Rectangle)
        {
            float left = (float)this.X;
            float top = (float)this.Y;
            float right = (float)(this.X + this.Width);
            float bottom = (float)(this.Y + this.Height);
            var rect = new SKRect(left, top, right, bottom);
            slideCanvas.DrawRect(rect, paint);
            var textFrame = (TextBox)this.TextBox!;
            textFrame.Draw(slideCanvas, left, top);
        }
    }
}