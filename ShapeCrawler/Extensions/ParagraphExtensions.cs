using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Extensions
{
    /// <summary>
    ///     Contains extension methods for <see cref="A.Paragraph" /> class object.
    /// </summary>
    public static class ParagraphExtensions
    {
        public static bool IsEmpty(this A.Paragraph aParagraph)
        {
            return !aParagraph.Descendants<A.Text>().Any();
        }
    }
}