using System;
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

    internal void UpdateIndentLevel(int level)
    {
        if (level is < 1 or > 9)
        {
            throw new ArgumentOutOfRangeException(nameof(level), level, "Indent level must be between 1 and 9.");
        }

        aParagraph.ParagraphProperties!.Level = new Int32Value(level - 1);
    }
}