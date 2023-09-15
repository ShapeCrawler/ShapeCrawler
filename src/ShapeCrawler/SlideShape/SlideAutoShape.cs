using System;
using System.Collections.Generic;
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideShape;

internal sealed class SlideAutoShape : IShape, IRemoveable
{
    private readonly P.Shape pShape;
    private readonly SimpleShape simpleShape;
    private readonly SlidePart sdkSlidePart;

    internal SlideAutoShape(
        SlidePart sdkSlidePart,
        P.Shape pShape) :
        this(
            sdkSlidePart,
            pShape,
            new SimpleShape(pShape),
            new SlideShapeOutline(sdkSlidePart, pShape.ShapeProperties!)
        )
    {
    }

    private SlideAutoShape(
        SlidePart sdkSlidePart,
        P.Shape pShape,
        SimpleShape simpleShape,
        SlideShapeOutline outline)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.pShape = pShape;
        this.simpleShape = simpleShape;
        this.Outline = outline;
    }

    public bool HasOutline => true;
    public IShapeOutline Outline { get; }
    public bool HasFill => this.simpleShape.HasFill;
    public SCShapeType ShapeType => SCShapeType.AutoShape;
    public IShapeFill Fill => this.simpleShape.Fill;
    public double Rotation => this.simpleShape.Rotation;
    public bool IsTextHolder => this.simpleShape.IsTextHolder;
    public ITextFrame TextFrame => this.simpleShape.TextFrame;

    public ITable AsTable() => this.simpleShape.AsTable();

    public IMediaShape AsMedia() => this.simpleShape.AsMedia();
    public bool IsPlaceholder => this.simpleShape.IsPlaceholder;
    public IPlaceholder Placeholder => this.simpleShape.Placeholder;

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
    
    internal IHtmlElement ToHtmlElement() => throw new NotImplementedException();

    #region Shape

    public int Width
    {
        get => this.simpleShape.Width;
        set => this.simpleShape.Width = value;
    }

    public int Height
    {
        get => this.simpleShape.Height;
        set => this.simpleShape.Height = value;
    }

    public int Id => this.simpleShape.Id;
    public string Name => this.simpleShape.Name;
    public bool Hidden => this.simpleShape.Hidden;
    public SCGeometry GeometryType => this.simpleShape.GeometryType;

    public string? CustomData
    {
        get => this.simpleShape.ParseCustomData();
        set => this.simpleShape.UpdateCustomData(value!);
    }

    public int X
    {
        get => this.simpleShape.X;
        set => this.simpleShape.X = value;
    }

    public int Y
    {
        get => this.simpleShape.Y;
        set => this.simpleShape.Y = value;
    }

    #endregion Shape

    internal void CopyTo(
        int id, 
        P.ShapeTree pShapeTree, 
        IEnumerable<string> existingShapeNames,
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