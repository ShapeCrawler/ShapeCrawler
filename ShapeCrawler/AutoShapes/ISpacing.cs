using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.AutoShapes;

/// <summary>
///     Represents a spacing of paragraph.
/// </summary>
public interface ISpacing
{
    /// <summary>
    ///     Gets the number of lines if Line Spacing specified in lines, otherwise <see langword="null"/>.
    /// </summary>
    double? LineSpacingLines { get; }

    /// <summary>
    ///     Gets the number of points if Line Spacing specified in points, otherwise <see langword="null"/>. 
    /// </summary>
    double? LineSpacingPoints { get; }
}

internal sealed class SCSpacing : ISpacing
{
    private readonly SCParagraph paragraph;
    private readonly A.Paragraph aParagraph;

    public SCSpacing(SCParagraph paragraph, A.Paragraph aParagraph)
    {
        this.paragraph = paragraph;
        this.aParagraph = aParagraph;
    }

    public double? LineSpacingLines => this.GetLineSpacingLines();

    public double? LineSpacingPoints => this.GetLineSpacingPoints();

    private double? GetLineSpacingLines()
    {
        var aLnSpc = this.aParagraph.ParagraphProperties!.LineSpacing;
        if (aLnSpc == null)
        {
            return 1;
        }

        var aSpcPct = aLnSpc.SpacingPercent;
        if (aSpcPct != null)
        {
            return aSpcPct.Val! * 1.0 / 100000;
        }

        return null;
    }

    private double? GetLineSpacingPoints()
    {
        var aLnSpc = this.aParagraph.ParagraphProperties!.LineSpacing;

        var aSpcPts = aLnSpc?.SpacingPoints;
        if (aSpcPts != null)
        {
            return aSpcPts.Val! * 1.0 / 100;
        }

        return null;
    }
}