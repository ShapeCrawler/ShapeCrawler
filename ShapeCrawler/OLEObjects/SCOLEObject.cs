using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using SkiaSharp;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.OLEObjects;

internal sealed class SCOLEObject : SlideSCShape, IOLEObject
{
    internal SCOLEObject(OpenXmlCompositeElement pShapeTreesChild, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide, SCGroupSCShape groupSCShape)
        : base(pShapeTreesChild, oneOfSlide, groupSCShape)
    {
    }

    public override SCGeometry GeometryType => SCGeometry.Rectangle;

    public override SCShapeType ShapeType => SCShapeType.OLEObject;

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }
}