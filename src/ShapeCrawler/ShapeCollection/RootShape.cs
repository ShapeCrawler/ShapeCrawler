using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideShape;
using ShapeCrawler.Texts;
using ShapeCrawler.Wrappers;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.ShapeCollection;

internal sealed class RootShape : CopyableShape, IRootShape
{
    private readonly IShape decoratedShape;
    private readonly P.Shape pShape;

    internal RootShape(
        TypedOpenXmlPart sdkTypedOpenXmlPart,
        P.Shape pShape,
        IShape decoratedShape)
        : base(sdkTypedOpenXmlPart, pShape)
    {
        this.decoratedShape = decoratedShape;
        this.pShape = pShape;
    }

    public void Duplicate()
    {
        var pShapeTree = (P.ShapeTree)this.pShape.Parent!;
        var autoShapes = new PShapeTreeWrap(pShapeTree);
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

    #region Decorated Shape

    public override ShapeType ShapeType => this.decoratedShape.ShapeType;
    public override bool HasOutline => this.decoratedShape.HasOutline;
    public override IShapeOutline Outline => this.decoratedShape.Outline;
    public override bool HasFill => this.decoratedShape.HasFill;
    public override IShapeFill Fill => this.decoratedShape.Fill;
    public override bool IsTextHolder => this.decoratedShape.IsTextHolder;
    public override ITextFrame TextFrame => this.decoratedShape.TextFrame;
    public override Geometry GeometryType => this.decoratedShape.GeometryType;

    public override int X
    {
        get => this.decoratedShape.X; 
        set => this.decoratedShape.X = value;
    }

    #endregion Decorated Shape
}