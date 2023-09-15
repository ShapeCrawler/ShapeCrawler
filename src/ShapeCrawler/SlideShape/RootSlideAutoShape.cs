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

namespace ShapeCrawler.SlideShape;

internal sealed class RootSlideAutoShape : IRootSlideShape
{
    private readonly SlidePart sdkSlidePart;
    private readonly IShape shape;
    private readonly P.Shape pShape;

    internal RootSlideAutoShape(
        SlidePart sdkSlidePart,
        P.Shape pShape,
        IShape shape)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.shape = shape;
        this.pShape = pShape;
    }

    #region Shape

    public bool HasOutline => this.shape.HasOutline;
    public IShapeOutline Outline => this.shape.Outline;
    public bool HasFill => this.shape.HasFill;

    public int Width
    {
        get => this.shape.Width;
        set => this.shape.Width = value;
    }

    public int Height
    {
        get => this.shape.Height;
        set => this.shape.Height = value;
    }

    public int Id => this.shape.Id;
    public string Name => this.shape.Name;
    public bool Hidden => this.shape.Hidden;
    public SCGeometry GeometryType => this.shape.GeometryType;
    public IShapeFill Fill => this.shape.Fill;
    public bool IsPlaceholder => this.shape.IsPlaceholder;
    public IPlaceholder Placeholder => this.shape.Placeholder;

    public string? CustomData
    {
        get => this.shape.CustomData;
        set => this.shape.CustomData = value;
    }

    public SCShapeType ShapeType => this.shape.ShapeType;

    public bool IsTextHolder => this.shape.IsTextHolder;

    public ITextFrame TextFrame => this.shape.TextFrame;

    public double Rotation => this.shape.Rotation;
    public ITable AsTable() => this.shape.AsTable();
    public IMediaShape AsMedia() => this.shape.AsMedia();

    public int X
    {
        get => this.shape.X;
        set => this.shape.X = value;
    }

    public int Y
    {
        get => this.shape.Y;
        set => this.shape.Y = value;
    }

    #endregion Shape

    public void Duplicate()
    {
        var pShapeTree = (P.ShapeTree)this.pShape.Parent!;
        var autoShapes = new Shapes(pShapeTree);
        autoShapes.Add(this.pShape);
    }

    void ICopyableShape.CopyTo(
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
}