using System.Linq;
using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Extensions
{
    /// <summary>
    /// Extension methods for <see cref="OpenXmlElement"/> instance.
    /// </summary>
    public static class XmlElementExtensions
    {
        /// <summary>
        /// Determines whether element is placeholder.
        /// </summary>
        public static bool IsPlaceholder(this OpenXmlElement openXmlElement)
        {
            return openXmlElement.Descendants<P.PlaceholderShape>().Any();
        }
    }
}
