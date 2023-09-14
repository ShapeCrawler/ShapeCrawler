using System;
using System.Collections.Generic;
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;
using Shape = ShapeCrawler.Shapes.Shape;

namespace ShapeCrawler.SlideShape;

internal sealed class SlideShape : IShape, IRemoveable
{
    private readonly P.Shape pShape;
    private readonly Shape shape;
    private readonly SlidePart sdkSlidePart;

    internal SlideShape(
        SlidePart sdkSlidePart,
        P.Shape pShape) :
        this(
            sdkSlidePart,
            pShape,
            new Shape(pShape),
            new SlideShapeOutline(sdkSlidePart, pShape.ShapeProperties!)
        )
    {
    }

    private SlideShape(
        SlidePart sdkSlidePart,
        P.Shape pShape,
        Shape shape,
        SlideShapeOutline outline)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.pShape = pShape;
        this.shape = shape;
        this.Outline = outline;
    }

    public bool HasOutline => true;
    public IShapeOutline Outline { get; }
    
    public SCShapeType ShapeType => SCShapeType.AutoShape;
    public IShapeFill Fill => this.ParseFill();

    public double Rotation { get; }
    
    public bool IsTextHolder => false;

    public ITextFrame TextFrame => new NullTextFrame();
    
    public ITable AsTable() =>
        throw new SCException(
            $"The shape is not a table. Use {nameof(IShape.ShapeType)} property to check if the shape is a table.");

    public IMediaShape AsMedia() =>
        throw new SCException(
            $"The shape is not a media shape. Use {nameof(IShape.ShapeType)} property to check if the shape is a media.");
    
    public bool IsPlaceholder => false;

    public IPlaceholder Placeholder => new NullPlaceholder();
    
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

    #region Shape
    
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
    public bool Hidden => this.shape.Hidden();

    public SCGeometry GeometryType => this.shape.GeometryType();

    public string? CustomData
    {
        get => this.shape.CustomData(); 
        set => this.shape.UpdateCustomData(value!);
    }
    
    public int X
    {
        get => this.shape.X();
        set => this.shape.UpdateX(value);
    }

    public int Y
    {
        get => this.shape.Y();
        set => this.shape.UpdateY(value);
    }
    
    #endregion Shape
    
    private SlideShapeFill ParseFill()
    {
        var useBgFill = this.pShape.UseBackgroundFill;
        return new SlideShapeFill(this.sdkSlidePart, this.pShape.GetFirstChild<P.ShapeProperties>() !, useBgFill);
    }

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

    void IRemoveable.Remove()
    {
        this.pShape.Remove();
    }
}