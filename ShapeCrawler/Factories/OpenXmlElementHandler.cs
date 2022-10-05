using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Factories
{
    /// <summary>
    ///     Represents a base class for a shape creator.
    /// </summary>
    internal abstract class OpenXmlElementHandler
    {
        /// <summary>
        ///     Gets or sets the next handler in the chain.
        /// </summary>
        internal OpenXmlElementHandler? Successor { get; set; }

        /// <summary>
        ///     Creates shape from child element of the <see cref="P.ShapeTree" /> element.
        /// </summary>
        internal abstract Shape? Create(OpenXmlCompositeElement compositeElementOfPShapeTree, SCSlide slide, SlideGroupShape groupShape);
    }
}