using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Texts;

// ReSharper disable once InconsistentNaming
internal sealed record SAParagraph
{
    private readonly A.Paragraph aParagraph;

    internal SAParagraph(A.Paragraph aParagraph)
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