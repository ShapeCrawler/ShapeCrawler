using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
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
}