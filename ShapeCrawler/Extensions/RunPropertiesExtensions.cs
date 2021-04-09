using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Extensions
{
    public static class RunPropertiesExtensions
    {
        /// <summary>
        ///     Gets instance of the <see cref="A.SolidFill"/> class.
        /// </summary>
        /// <returns><see cref="A.SolidFill"/> instance or NULL.</returns>
        public static A.SolidFill SolidFill(this A.RunProperties aRunProperties)
        {
            return aRunProperties.GetFirstChild<A.SolidFill>();
        }
    }
}