using System.Collections.Generic;
using DocumentFormat.OpenXml;
using ShapeCrawler.Settings;
using ShapeCrawler.Shapes;

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
            SCSlide slide)
            : base(slide, pShapeTreeChild)
        {
            this.Context = spContext;
            this.Shapes = groupedShapes;
        }

        #endregion Constructors

        #region Public Properties

        public IReadOnlyCollection<IShape> Shapes { get; }

        #endregion Properties
    }
}