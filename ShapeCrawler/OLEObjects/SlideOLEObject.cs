using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a slide.
    /// </summary>
    internal class SlideOLEObject : SlideShape, IOLEObject // TODO: make it internal
    {
        internal SlideOLEObject(
            OpenXmlCompositeElement pShapeTreesChild,
            SCSlide parentSlideLayoutInternal,
            SlideGroupShape groupShape)
            : base(pShapeTreesChild, parentSlideLayoutInternal, groupShape)
        {
        }

        #region Public Properties

        public override SCGeometry GeometryType => SCGeometry.Rectangle;

        public SCShapeType ShapeType => SCShapeType.OLEObject;

        #endregion Public Properties
    }
}