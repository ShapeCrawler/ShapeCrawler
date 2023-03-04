using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.Shapes;
using SkiaSharp;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.OLEObjects;

internal sealed class SCOLEObject : SCShape, IOLEObject
{
    internal SCOLEObject(
        OpenXmlCompositeElement pShapeTreesChild, 
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> parentSlideObject,
        OneOf<ShapeCollection, SCGroupShape> parentShapeCollection)
        : base(pShapeTreesChild, parentSlideObject, parentShapeCollection)
    {
    }

    public override SCGeometry GeometryType => SCGeometry.Rectangle;

    public override SCShapeType ShapeType => SCShapeType.OLEObject;

    internal override void Draw(SKCanvas canvas)
    {
        throw new System.NotImplementedException();
    }
}