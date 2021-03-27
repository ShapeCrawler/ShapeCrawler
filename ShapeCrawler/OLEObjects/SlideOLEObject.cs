using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using ShapeCrawler.OLEObjects;
using ShapeCrawler.Settings;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape on a slide.
    /// </summary>
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class SlideOLEObject : SlideShape, IOLEObject //Make internal
    {
        #region Constructors

        internal SlideOLEObject(
            OpenXmlCompositeElement shapeTreeChild,
            ShapeContext spContext,
            SCSlide slide) : base(slide, shapeTreeChild)
        {
            ShapeTreeChild = shapeTreeChild;
            Context = spContext;
        }

        #endregion Constructors

        #region Public Properties

        public override GeometryType GeometryType => GeometryType.Rectangle;

        #endregion Public Properties

        #region Fields

        internal ShapeContext Context;

        internal OpenXmlCompositeElement ShapeTreeChild { get; }

        #endregion Fields
    }
}