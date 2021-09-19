using System.Collections.Generic;
using DocumentFormat.OpenXml;
using ShapeCrawler.Factories;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a group shape on a Slide.
    /// </summary>
    internal class SlideGroupShape : SlideShape, IGroupShape
    {
        internal ShapeContext Context;

        #region Constructors

        internal SlideGroupShape(
            ShapeContext spContext,
            List<IShape> groupedShapes,
            OpenXmlCompositeElement pShapeTreeChild,
            SCSlide slide) : base(slide, pShapeTreeChild)
        {
            Context = spContext;
            Shapes = groupedShapes;
        }

        #endregion Constructors

        #region Public Properties

        public IReadOnlyCollection<IShape> Shapes { get; }

        #endregion Properties
    }
}