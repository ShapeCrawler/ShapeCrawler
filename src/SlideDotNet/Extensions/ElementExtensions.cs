using System.Linq;
using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace SlideDotNet.Extensions
{
    /// <summary>
    /// Extension methods for <see cref="OpenXmlElement"/> instance.
    /// </summary>
    public static class ElementExtensions
    {
        /// <summary>
        /// Determines whether element is placeholder.
        /// </summary>
        public static bool IsPlaceholder(this OpenXmlElement xmlElement)
        {
            return xmlElement.Descendants<P.PlaceholderShape>().Any();
        }
    }
}
