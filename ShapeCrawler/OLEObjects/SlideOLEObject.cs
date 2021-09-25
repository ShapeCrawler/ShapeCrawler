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
        internal ShapeContext Context;

        internal OpenXmlCompositeElement PShapeTreesChild { get; } // TODO: is it extra property?

        #region Constructors

        internal SlideOLEObject(
            SCSlide parentSlide,
            OpenXmlCompositeElement pShapeTreesChild,
            ShapeContext spContext)
            : base(parentSlide, pShapeTreesChild)
        {
            //this.PShapeTreesChild = pShapeTreesChild;
            this.Context = spContext;
        }

        internal SlideOLEObject(
            SCSlide parentSlide,
            OpenXmlCompositeElement pShapeTreesChild, 
            ShapeContext spContext, 
            SlideGroupShape parentGroupShape)
            : base(parentSlide, pShapeTreesChild, parentGroupShape)
        {
            //this.PShapeTreesChild = pShapeTreesChild;
            this.Context = spContext;
        }

        #endregion Constructors

        #region Public Properties

        public override GeometryType GeometryType => GeometryType.Rectangle;

        #endregion Public Properties
    }
}