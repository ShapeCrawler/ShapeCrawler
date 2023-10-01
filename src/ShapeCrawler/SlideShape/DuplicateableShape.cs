using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideShape;

internal sealed class DuplicateableShape : CopyableShape, IDuplicateableShape
{
    private readonly IShape shape;
    private readonly P.Shape pShape;

    internal DuplicateableShape(
        TypedOpenXmlPart sdkTypedOpenXmlPart,
        P.Shape pShape,
        IShape shape)
        : base(sdkTypedOpenXmlPart, pShape)
    {
        this.shape = shape;
        this.pShape = pShape;
    }

    public void Duplicate()
    {
        var pShapeTree = (P.ShapeTree)this.pShape.Parent!;
        var autoShapes = new Shapes(pShapeTree);
        autoShapes.Add(this.pShape);
    }

    internal void Draw(SKCanvas slideCanvas)
    {
        var skColorOutline = SKColor.Parse(this.Outline.HexColor);

        using var paint = new SKPaint
        {
            Color = skColorOutline,
            IsAntialias = true,
            StrokeWidth = UnitConverter.PointToPixel(this.Outline.Weight),
            Style = SKPaintStyle.Stroke
        };

        if (this.GeometryType == Geometry.Rectangle)
        {
            float left = this.X;
            float top = this.Y;
            float right = this.X + this.Width;
            float bottom = this.Y + this.Height;
            var rect = new SKRect(left, top, right, bottom);
            slideCanvas.DrawRect(rect, paint);
            var textFrame = (TextFrame)this.TextFrame!;
            textFrame.Draw(slideCanvas, left, this.Y);
        }
    }

    #region Shape

    public override ShapeType ShapeType => this.shape.ShapeType;
    public override bool HasOutline => this.shape.HasOutline;
    public override IShapeOutline Outline => this.shape.Outline;
    public override bool HasFill => this.shape.HasFill;
    public override IShapeFill Fill => this.shape.Fill;

    #endregion Shape
}