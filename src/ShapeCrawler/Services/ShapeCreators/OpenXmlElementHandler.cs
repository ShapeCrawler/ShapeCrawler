using DocumentFormat.OpenXml;
using SlideDotNet.Models.SlideComponents;

namespace SlideDotNet.Services.ShapeCreators
{
    /// <summary>
    /// Represents a base class for shape creators.
    /// </summary>
    public abstract class OpenXmlElementHandler
    {
        /// <summary>
        /// Gets or sets the next handler in the chain.
        /// </summary>
        public OpenXmlElementHandler Successor { get; set; }
        
        /// <summary>
        /// Creates <see cref="ShapeEx"/> instance from specified SDK element or passes it to next handler.
        /// </summary>
        /// <param name="sdkElement"></param>
        /// <returns></returns>
        public abstract ShapeEx Create(OpenXmlElement sdkElement);
    }
}
