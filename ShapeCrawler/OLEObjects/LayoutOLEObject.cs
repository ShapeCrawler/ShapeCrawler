using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable InconsistentNaming
namespace ShapeCrawler.OLEObjects;

internal class LayoutOLEObject : LayoutShape, IShape
{
    internal LayoutOLEObject(SCSlideLayout slideLayoutInternal, P.GraphicFrame pGraphicFrame)
        : base(slideLayoutInternal, pGraphicFrame)
    {
    }

    public override SCShapeType ShapeType => SCShapeType.OLEObject;

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }
}