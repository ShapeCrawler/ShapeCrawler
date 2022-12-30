using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using SkiaSharp;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.OLEObjects;

internal sealed class SlideOLEObject : SlideShape, IOLEObject
{
    internal SlideOLEObject(OpenXmlCompositeElement pShapeTreesChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide, SCGroupShape groupShape)
        : base(pShapeTreesChild, oneOfSlide, groupShape)
    {
    }

    public override SCGeometry GeometryType => SCGeometry.Rectangle;

    public override SCShapeType ShapeType => SCShapeType.OLEObject;

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }
}