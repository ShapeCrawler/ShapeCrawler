using A = DocumentFormat.OpenXml.Drawing;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

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

internal sealed class Spacing(Paragraph paragraph, A.Paragraph aParagraph): ISpacing
{
    private readonly Paragraph paragraph = paragraph;

    public double? LineSpacingLines => this.GetLineSpacingLines();

    public double? LineSpacingPoints => this.GetLineSpacingPoints();

    private double? GetLineSpacingLines()
    {
        var aLnSpc = aParagraph.ParagraphProperties!.LineSpacing;
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
        var aLnSpc = aParagraph.ParagraphProperties!.LineSpacing;

        var aSpcPts = aLnSpc?.SpacingPoints;
        if (aSpcPts != null)
        {
            return aSpcPts.Val! * 1.0 / 100;
        }

        return null;
    }
}