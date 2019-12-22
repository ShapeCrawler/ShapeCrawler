using P = DocumentFormat.OpenXml.Presentation;

namespace PptxXML.Services.Placeholder
{
    /// <summary>
    /// Represents a data of a placeholder.
    /// </summary>
    public class PlaceholderData
    {
        /// <summary>
        /// Gets or sets X-coordinate's value.
        /// </summary>
        public long X { get; set; }

        /// <summary>
        /// Gets or sets Y-coordinate's value.
        /// </summary>
        public long Y { get; set; }

        /// <summary>
        /// Gets or sets width value.
        /// </summary>
        public long Width { get; set; }

        /// <summary>
        /// Gets or sets height value.
        /// </summary>
        public long Height { get; set; }

        /// <summary>
        /// Gets or set geometry form code.
        /// </summary>
        public int GeometryCode { get; set; }

        /// <summary>
        /// Gets or sets the instance of the <see cref="P.ShapeProperties"/> class.
        /// </summary>
        public P.ShapeProperties ShapeProperties { get; set; }
    }
}
