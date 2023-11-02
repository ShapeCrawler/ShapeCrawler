using System;
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.ShapeCollection;

internal sealed class AutoShape : CopyableShape
{
    private readonly P.Shape pShape;

    internal AutoShape(
        TypedOpenXmlPart sdkTypedOpenXmlPart,
        P.Shape pShape,
        TextFrame textFrame)
        : this(sdkTypedOpenXmlPart, pShape)
    {
        this.IsTextHolder = true;
        this.TextFrame = textFrame;
    }

    internal AutoShape(
        TypedOpenXmlPart sdkTypedOpenXmlPart,
        P.Shape pShape)
        : base(sdkTypedOpenXmlPart, pShape)
    {
        this.pShape = pShape;
        this.Outline = new SlideShapeOutline(this.sdkTypedOpenXmlPart, pShape.Descendants<P.ShapeProperties>().First());
        this.Fill = new ShapeFill(this.sdkTypedOpenXmlPart, pShape.Descendants<P.ShapeProperties>().First());
    }

    public override bool HasOutline => true;
    public override IShapeOutline Outline { get; }
    public override bool HasFill => true;
    public override IShapeFill Fill { get; }
    public override ShapeType ShapeType => ShapeType.AutoShape;

    public override Geometry GeometryType
    {
        get
        {
            var spPr = this.pShapeTreeElement.Descendants<P.ShapeProperties>().First();
            var aPresetGeometry = spPr.GetFirstChild<A.PresetGeometry>();

            if (aPresetGeometry == null)
            {
                if (spPr.OfType<A.CustomGeometry>().Any())
                {
                    return Geometry.Custom;
                }
            }
            else
            {
                var name = aPresetGeometry.Preset!.Value.ToString();
                Enum.TryParse(name, true, out Geometry geometryType);
                return geometryType;    
            }
            
            return Geometry.Rectangle;
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

    internal IHtmlElement ToHtmlElement() => throw new NotImplementedException();

    public override bool Removeable => true;
    public override void Remove() => this.pShape.Remove();
}