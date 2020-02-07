using DocumentFormat.OpenXml;
using SlideXML.Enums;

namespace SlideXML.Services
{
    /// <summary>
    /// Represents a parsed candidate element.
    /// </summary>
    public class ElementCandidate
    {
        /// <summary>
        /// Gets or sets corresponding element type.
        /// </summary>
        public ElementType ElementType { get; set; }

        /// <summary>
        /// Gets or sets instance of <see cref="OpenXmlCompositeElement"/>.
        /// </summary>
        public OpenXmlCompositeElement CompositeElement;
    }
}