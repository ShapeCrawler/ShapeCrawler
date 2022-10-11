using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using OneOf;

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

    #region Public Properties

    public override SCGeometry GeometryType => SCGeometry.Rectangle;

    public override SCShapeType ShapeType => SCShapeType.OLEObject;

    #endregion Public Properties
}