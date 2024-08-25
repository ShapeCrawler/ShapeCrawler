using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.ShapeCollection;

internal sealed class RootShape : CopyableShape, IRootShape
{
    private readonly IShape decoratedShape;
    private readonly P.Shape pShape;

    internal RootShape(
        OpenXmlPart sdkTypedOpenXmlPart,
        P.Shape pShape,
        IShape decoratedShape)
        : base(sdkTypedOpenXmlPart, pShape)
    {
        this.decoratedShape = decoratedShape;
        this.pShape = pShape;
    }

    #region Decorated Shape

    public override ShapeType ShapeType => this.decoratedShape.ShapeType;
    
    public override bool HasOutline => this.decoratedShape.HasOutline;
    
    public override IShapeOutline Outline => this.decoratedShape.Outline;
    
    public override bool HasFill => this.decoratedShape.HasFill;
    
    public override IShapeFill Fill => this.decoratedShape.Fill;
    
    public override bool IsTextHolder => this.decoratedShape.IsTextHolder;
    
    public override ITextBox TextBox => this.decoratedShape.TextBox;
    
    public override Geometry GeometryType => this.decoratedShape.GeometryType;

    public override decimal X
    {
        get => this.decoratedShape.X; 
        set => this.decoratedShape.X = value;
    }

    #endregion Decorated Shape
    
    public void Duplicate()
    {
        var pShapeTree = (P.ShapeTree)this.pShape.Parent!;
        var autoShapes = new WrappedPShapeTree(pShapeTree);
        autoShapes.Add(this.pShape);
    }

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
            float left = (float)(this.X);
            float top = (float)(this.Y);
            float right = (float)(this.X + this.Width);
            float bottom = (float)(this.Y + this.Height);
            var rect = new SKRect(left, top, right, bottom);
            slideCanvas.DrawRect(rect, paint);
            var textFrame = (TextFrame)this.TextBox!;
            textFrame.Draw(slideCanvas, left, top);
        }
    }
}