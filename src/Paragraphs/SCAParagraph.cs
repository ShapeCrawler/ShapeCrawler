using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable InconsistentNaming
namespace ShapeCrawler.Paragraphs;

internal sealed class SCAParagraph(A.Paragraph aParagraph)
{
    internal int GetIndentLevel()
    {
        var level = aParagraph.ParagraphProperties!.Level;
        if (level is null)
        {
            return 1; // default indent level
        }

        return level + 1;
    }

    internal void UpdateIndentLevel(int level) => aParagraph.ParagraphProperties!.Level = new Int32Value(level - 1);
}