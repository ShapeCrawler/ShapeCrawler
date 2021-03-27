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
        public OpenXmlElementHandler Successor { get; set; }

        /// <summary>
        ///     Creates shape from child element of the <see cref="P.ShapeTree" /> element.
        /// </summary>
        public abstract IShape Create(OpenXmlCompositeElement pShapeTreeChild, SCSlide slide);
    }
}