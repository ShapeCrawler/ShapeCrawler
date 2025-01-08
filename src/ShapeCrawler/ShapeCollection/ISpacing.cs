using System;
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
    
    /// <summary>
    ///    Gets or sets the number of points before the paragraph, otherwise <see langword="null"/>.
    /// </summary>
    double BeforeSpacingPoints { get; set; }
    
    /// <summary>
    ///   Gets or sets the number of points after the paragraph, otherwise <see langword="null"/>.
    /// </summary>
    double AfterSpacingPoints { get; set; }
}

internal sealed class Spacing(A.Paragraph aParagraph): ISpacing
{
    public double? LineSpacingLines => this.GetLineSpacingLines();

    public double? LineSpacingPoints => this.GetLineSpacingPoints();

    public double BeforeSpacingPoints
    {
        get => this.GetBeforeSpacingPoints();
        set => this.SetBeforeSpacingPoints(value);
    }

    public double AfterSpacingPoints
    {
        get => this.GetAfterSpacingPoints();
        set => this.SetAfterSpacingPoints(value);
    }
    
    private static double ConvertHundredsOfPointsToPoints(int hundredsOfPoints)
    {
        return hundredsOfPoints * 1.0 / 100;
    }
    
    private static int ConvertPointsToHundredsOfPoints(double points)
    {
        return (int)Math.Round(points * 100);
    }
    
    private double GetBeforeSpacingPoints()
    {
        var aSpcBef = aParagraph.ParagraphProperties?.SpaceBefore?.SpacingPoints?.Val;

        return aSpcBef != null ? ConvertHundredsOfPointsToPoints(aSpcBef) : 0;
    }
    
    private void SetBeforeSpacingPoints(double points)
    {
        var aSpcBef = aParagraph.ParagraphProperties;
        aSpcBef ??= new A.ParagraphProperties();
        aSpcBef.SpaceBefore ??= new A.SpaceBefore();
        aSpcBef.SpaceBefore.SpacingPoints ??= new A.SpacingPoints();
        
        var hundredsOfPoints = ConvertPointsToHundredsOfPoints(points);
        
        if (hundredsOfPoints == 0)
        {
            aSpcBef.SpaceBefore = null;
        }
        else
        {
            aSpcBef.SpaceBefore.SpacingPoints.Val = hundredsOfPoints;
        }
    }
    
    private double GetAfterSpacingPoints()
    {
        var aSpcAft = aParagraph.ParagraphProperties?.SpaceAfter?.SpacingPoints?.Val;
        
        return aSpcAft != null ? ConvertHundredsOfPointsToPoints(aSpcAft) : 0;
    }
    
    private void SetAfterSpacingPoints(double points)
    {
        var aSpcAft = aParagraph.ParagraphProperties;
        aSpcAft ??= new A.ParagraphProperties();
        aSpcAft.SpaceAfter ??= new A.SpaceAfter();
        aSpcAft.SpaceAfter.SpacingPoints ??= new A.SpacingPoints();

        var hundredsOfPoints = ConvertPointsToHundredsOfPoints(points);
        
        if (hundredsOfPoints == 0)
        {
            aSpcAft.SpaceAfter = null;
        }
        else
        {
            aSpcAft.SpaceAfter.SpacingPoints.Val = hundredsOfPoints;
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

    private double? GetLineSpacingPoints()
    {
        var aLnSpc = aParagraph.ParagraphProperties!.LineSpacing?.SpacingPoints;

        if (aLnSpc != null)
        {
            return ConvertHundredsOfPointsToPoints(aLnSpc.Val!);
        }

        return null;
    }
}