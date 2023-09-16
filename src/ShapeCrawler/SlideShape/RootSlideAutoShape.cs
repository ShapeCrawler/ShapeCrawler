using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideShape;

internal sealed class RootSlideAutoShape : IRootSlideAutoShape
{
    private readonly SlideAutoShape slideAutoShape;
    private readonly SlidePart sdkSlidePart;
    private readonly P.Shape pShape;

    internal RootSlideAutoShape(
        SlidePart sdkSlidePart,
        P.Shape pShape,
        SlideAutoShape slideAutoShape)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.slideAutoShape = slideAutoShape;
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

        if (this.GeometryType == SCGeometry.Rectangle)
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

    #region ICopyableShape

    public bool HasOutline => this.slideAutoShape.HasOutline;
    public IShapeOutline Outline => this.slideAutoShape.Outline;
    public bool HasFill => this.slideAutoShape.HasFill;

    public int Width
    {
        get => this.slideAutoShape.Width;
        set => this.slideAutoShape.Width = value;
    }

    public int Height
    {
        get => this.slideAutoShape.Height;
        set => this.slideAutoShape.Height = value;
    }

    public int Id => this.slideAutoShape.Id;
    public string Name => this.slideAutoShape.Name;
    public bool Hidden => this.slideAutoShape.Hidden;
    public SCGeometry GeometryType => this.slideAutoShape.GeometryType;
    public IShapeFill Fill => this.slideAutoShape.Fill;
    public bool IsPlaceholder => this.slideAutoShape.IsPlaceholder;
    public IPlaceholder Placeholder => this.slideAutoShape.Placeholder;

    public string? CustomData
    {
        get => this.slideAutoShape.CustomData;
        set => this.slideAutoShape.CustomData = value;
    }

    public SCShapeType ShapeType => this.slideAutoShape.ShapeType;

    public bool IsTextHolder => this.slideAutoShape.IsTextHolder;

    public ITextFrame TextFrame => this.slideAutoShape.TextFrame;

    public double Rotation => this.slideAutoShape.Rotation;
    public ITable AsTable() => this.slideAutoShape.AsTable();
    public IMediaShape AsMedia() => this.slideAutoShape.AsMedia();
    public void CopyTo(int id, P.ShapeTree pShapeTree, IEnumerable<string> existingShapeNames, SlidePart targetSdkSlidePart)
    {
        this.slideAutoShape.CopyTo(id, pShapeTree, existingShapeNames, targetSdkSlidePart);
    }

    public int X
    {
        get => this.slideAutoShape.X;
        set => this.slideAutoShape.X = value;
    }

    public int Y
    {
        get => this.slideAutoShape.Y;
        set => this.slideAutoShape.Y = value;
    }

    
    #endregion ICopyableShape
}