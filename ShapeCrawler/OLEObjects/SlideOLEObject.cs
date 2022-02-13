using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using ShapeCrawler.OLEObjects;
using ShapeCrawler.Settings;
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

        public override GeometryType GeometryType => GeometryType.Rectangle;

        public ShapeType ShapeType => ShapeType.OLEObject;

        #endregion Public Properties
    }
}