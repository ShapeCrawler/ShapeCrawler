using DocumentFormat.OpenXml;
using ShapeCrawler.Models;

namespace ShapeCrawler.Factories.ShapeCreators
{
    /// <summary>
    /// Represents a base class for shape creators.
    /// </summary>
    internal abstract class OpenXmlElementHandler
    {
        /// <summary>
        /// Gets or sets the next handler in the chain.
        /// </summary>
        public OpenXmlElementHandler Successor { get; set; }
        
        /// <summary>
        /// Creates <see cref="GroupShapeSc"/> instance from specified SDK element or passes it to next handler.
        /// </summary>
        public abstract IShape Create(OpenXmlCompositeElement shapeTreeSource, SlideSc slide);
    }
}
