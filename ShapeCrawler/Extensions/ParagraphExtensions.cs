using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Extensions;

internal static class ParagraphExtensions
{
    internal static bool IsEmpty(this A.Paragraph aParagraph)
    {
        return !aParagraph.Descendants<A.Text>().Any();
    }
}