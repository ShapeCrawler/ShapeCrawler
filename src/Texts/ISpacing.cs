using ShapeCrawler.Units;
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
    ///     Gets the number of spaces in lines appears between lines of text. Returns <see langword="null"/> if the spaces are not specified in lines.
    /// </summary>
    double? LineSpacingLines { get; }

    /// <summary>
    ///     Gets the number of spaces in points appears between lines of text. Returns <see langword="null"/> if the spaces are not specified in points. 
    /// </summary>
    decimal? LineSpacingPoints { get; }

    /// <summary>
    ///    Gets or sets the number of spaces in points before the paragraph.
    /// </summary>
    decimal BeforeSpacingPoints { get; set; }

    /// <summary>
    ///    Gets or sets the number of spaces in points after the paragraph.
    /// </summary>
    decimal AfterSpacingPoints { get; set; }
}

internal sealed class Spacing(A.Paragraph aParagraph): ISpacing
{
    public double? LineSpacingLines => this.GetLineSpacingLines();

    public decimal? LineSpacingPoints
    {
        get
        {
            var aLnSpc = aParagraph.ParagraphProperties!.LineSpacing?.SpacingPoints;
            if (aLnSpc is not null)
            {
                return aLnSpc.Val! / 100m;
            }

            return null;
        }
    }

    public decimal BeforeSpacingPoints
    {
        get
        {
            var hundredsOfPoints = aParagraph.ParagraphProperties?.SpaceBefore?.SpacingPoints?.Val;
            if (hundredsOfPoints is null)
            {
                return 0;
            }

            return hundredsOfPoints / 100m;
        }
        set => this.SetBeforeSpacingPoints(value);
    }

    public decimal AfterSpacingPoints
    {
        get
        {
            var hundredsOfPoints = aParagraph.ParagraphProperties?.SpaceAfter?.SpacingPoints?.Val;
            if (hundredsOfPoints is null)
            {
                return 0;
            }

            return hundredsOfPoints / 100m;
        }

        set
        {
            var aSpcAft = aParagraph.ParagraphProperties;
            aSpcAft ??= new A.ParagraphProperties();
            aSpcAft.SpaceAfter ??= new A.SpaceAfter();
            aSpcAft.SpaceAfter.SpacingPoints ??= new A.SpacingPoints();

            var hundredsOfPoints = new Points(value).AsHundredPoints();

            if (hundredsOfPoints == 0)
            {
                aSpcAft.SpaceAfter = null;
            }
            else
            {
                aSpcAft.SpaceAfter.SpacingPoints.Val = hundredsOfPoints;
            }
        }
    }

    private void SetBeforeSpacingPoints(decimal points)
    {
        var aSpcBef = aParagraph.ParagraphProperties;
        aSpcBef ??= new A.ParagraphProperties();
        aSpcBef.SpaceBefore ??= new A.SpaceBefore();
        aSpcBef.SpaceBefore.SpacingPoints ??= new A.SpacingPoints();

        var hundredsOfPoints = new Points(points).AsHundredPoints();

        if (hundredsOfPoints == 0)
        {
            aSpcBef.SpaceBefore = null;
        }
        else
        {
            aSpcBef.SpaceBefore.SpacingPoints.Val = hundredsOfPoints;
        }
    }

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
}