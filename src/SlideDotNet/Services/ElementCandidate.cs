using DocumentFormat.OpenXml;
using SlideDotNet.Enums;

namespace SlideDotNet.Services
{
    /// <summary>
    /// Represents a parsed candidate element.
    /// </summary>
    public class ElementCandidate //TODO: [optimize] consider the possibility to use struct instead of class
    {
        /// <summary>
        /// Gets or sets corresponding element type.
        /// </summary>
        public ShapeContentType ElementType { get; set; }

        /// <summary>
        /// Gets or sets instance of <see cref="OpenXmlCompositeElement"/>.
        /// </summary>
        public OpenXmlCompositeElement XmlElement;
    }
}