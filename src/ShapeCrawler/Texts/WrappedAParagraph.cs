using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

internal sealed record WrappedAParagraph
{
    private readonly A.Paragraph aParagraph;

    internal WrappedAParagraph(A.Paragraph aParagraph)
    {
        this.aParagraph = aParagraph;
    }

    internal int IndentLevel()
    {
        var level = this.aParagraph.ParagraphProperties!.Level;
        if (level is null)
        {
            return 1; // default indent level
        }

        return level + 1;
    }
    
    internal void UpdateIndentLevel(int level) => this.aParagraph.ParagraphProperties!.Level = new Int32Value(level - 1);
}