using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Extensions;

internal static class ParagraphExtensions
{
    internal static bool IsEmpty(this A.Paragraph aParagraph)
    {
        // Consider paragraph empty if it has no text runs or all text runs are empty strings
        return !aParagraph.Descendants<A.Text>().Any(t => !string.IsNullOrEmpty(t.Text));
    }
}