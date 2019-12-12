using P = DocumentFormat.OpenXml.Presentation;

namespace PptxXML.Services
{
    /// <summary>
    /// Represents data of a placeholder.
    /// </summary>
    public class PlaceholderData
    {
        public long X { get; set; }

        public long Y { get; set; }

        public long Width { get; set; }

        public long Height { get; set; }

        /// <summary>
        /// Gets or set geometry form code.
        /// </summary>
        public int GeometryCode { get; set; }

        public P.ShapeProperties ShapeProperties { get; set; }
    }
}
