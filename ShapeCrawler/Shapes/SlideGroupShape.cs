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
        internal SlideGroupShape(
            List<IShape> groupedShapes,
            OpenXmlCompositeElement pShapeTreeChild,
            SCSlide slide)
            : base(slide, pShapeTreeChild)
        {
            this.Shapes = groupedShapes;
        }

        public IReadOnlyCollection<IShape> Shapes { get; }
    }
}