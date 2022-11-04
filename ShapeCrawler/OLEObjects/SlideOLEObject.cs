using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using OneOf;
using SkiaSharp;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.OLEObjects;

/// <summary>
///     Represents a shape on a slide.
/// </summary>
internal class SlideOLEObject : SlideShape, IOLEObject
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