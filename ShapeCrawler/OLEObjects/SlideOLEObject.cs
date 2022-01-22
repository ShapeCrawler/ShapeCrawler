using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using ShapeCrawler.OLEObjects;
using ShapeCrawler.Settings;

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
            SlideGroupShape parentGroupShape)
            : base(pShapeTreesChild, parentSlideLayoutInternal, parentGroupShape)
        {
        }

        #region Public Properties

        public override GeometryType GeometryType => GeometryType.Rectangle;

        #endregion Public Properties
    }
}