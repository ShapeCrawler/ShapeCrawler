using System;
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideShape;

internal sealed class SlideAutoShape : CopyableShape, IShape, IRemoveable
{
    private readonly P.Shape pShape;
    private readonly SlidePart sdkSlidePart;

    internal SlideAutoShape(
        SlidePart sdkSlidePart,
        P.Shape pShape)
        : base(pShape)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.pShape = pShape;
        this.Outline = new SlideShapeOutline(sdkSlidePart, pShape.Descendants<P.ShapeProperties>().First());
        this.Fill = new SlideShapeFill(sdkSlidePart, pShape.Descendants<P.ShapeProperties>().First(), false);
    }

    public override bool HasOutline => true;
    public override IShapeOutline Outline { get; }
    public override bool HasFill => true;
    public override IShapeFill Fill { get; }
    public override SCShapeType ShapeType => SCShapeType.AutoShape;

    public override SCGeometry GeometryType
    {
        get
        {
            var spPr = this.pShapeTreeElement.Descendants<P.ShapeProperties>().First();
            var aPresetGeometry = spPr.GetFirstChild<A.PresetGeometry>();
            
            if (aPresetGeometry == null) // Placeholder can have transform on the slide, without having geometry
            {
                return SCGeometry.Custom;
            }

            var name = aPresetGeometry.Preset!.Value.ToString();
            Enum.TryParse(name, true, out SCGeometry geometryType);
            return geometryType;
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

    internal IHtmlElement ToHtmlElement() => throw new NotImplementedException();

    void IRemoveable.Remove() => this.pShape.Remove();
}