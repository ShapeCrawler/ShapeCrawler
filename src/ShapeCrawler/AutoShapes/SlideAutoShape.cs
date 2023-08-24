using System;
using System.Collections.Generic;
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes;

internal sealed record SlideAutoShape : ISlideAutoShape
{
    private readonly P.Shape pShape;
    private readonly Shape shape;
    private readonly Lazy<SlideAutoShapeFill> autoShapeFill;
    private readonly SlidePart sdkSlidePart;

    private event Action Duplicated;

    internal SlideAutoShape(
        SlidePart sdkSlidePart, 
        P.Shape pShape, 
        Shape shape, 
        SlideShapeOutline outline, 
        Action duplicatedHandler)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.pShape = pShape;
        this.shape = shape;
        this.Outline = outline;
        this.autoShapeFill = new Lazy<SlideAutoShapeFill>(this.ParseFill);
        this.Duplicated += duplicatedHandler;
    }

    public IShapeOutline Outline { get; }

    public int Width
    {
        get => this.shape.Width();
        set => this.shape.UpdateWidth(value);
    }

    public int Height
    {
        get => this.shape.Height();
        set => this.shape.UpdateHeight(value);
    }

    public int Id => this.shape.Id();
    public string Name => this.shape.Name();
    public bool Hidden { get; }
    public bool IsPlaceholder() => false;

    public IPlaceholder Placeholder => new NullPlaceholder();

    public SCGeometry GeometryType { get; }
    public string? CustomData { get; set; }
    public SCShapeType ShapeType => SCShapeType.AutoShape;
    public IAutoShape AsAutoShape() => this;

    public IShapeFill Fill => this.autoShapeFill.Value;

    public ITextFrame TextFrame => new NullTextFrame();

    public bool IsTextHolder() => false;

    public double Rotation { get; }

    public void Duplicate()
    {
        var pShapeTree = (P.ShapeTree)pShape.Parent!;
        var autoShapes = new SlideAutoShapes(pShapeTree);
        autoShapes.Add(this.pShape);

        this.Duplicated.Invoke();
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

    internal string ToJson()
    {
        throw new NotImplementedException();
    }

    internal IHtmlElement ToHtmlElement()
    {
        throw new NotImplementedException();
    }

    private SlideAutoShapeFill ParseFill()
    {
        var useBgFill = pShape.UseBackgroundFill;
        return new SlideAutoShapeFill(this.sdkSlidePart, this.pShape.GetFirstChild<P.ShapeProperties>() !, useBgFill);
    }

    public int X { get; set; }
    public int Y { get; set; }

    internal void CopyTo(int id, P.ShapeTree pShapeTree, IEnumerable<string> existingShapeNames,
        SlidePart targetSdkSlidePart)
    {
        var copy = this.pShape.CloneNode(true);
        copy.GetNonVisualDrawingProperties().Id = new UInt32Value((uint)id);
        pShapeTree.AppendChild(copy);
        var copyName = copy.GetNonVisualDrawingProperties().Name!.Value!;
        if (existingShapeNames.Any(existingShapeName => existingShapeName == copyName))
        {
            var currentShapeCollectionSuffixes = existingShapeNames
                .Where(c => c.StartsWith(copyName, StringComparison.InvariantCulture))
                .Select(c => c.Substring(copyName.Length))
                .ToArray();

            // We will try to check numeric suffixes only.
            var numericSuffixes = new List<int>();

            foreach (var currentSuffix in currentShapeCollectionSuffixes)
            {
                if (int.TryParse(currentSuffix, out var numericSuffix))
                {
                    numericSuffixes.Add(numericSuffix);
                }
            }

            numericSuffixes.Sort();
            var lastSuffix = numericSuffixes.LastOrDefault() + 1;
            copy.GetNonVisualDrawingProperties().Name = copyName + " " + lastSuffix;
        }
    }
}