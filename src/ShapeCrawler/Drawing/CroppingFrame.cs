using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Exceptions;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing;

/// <summary>
///     The portion within an image which is visible. 
/// </summary>
/// <remarks>
///     Values are a fraction of the whole image, e.g. `0.25` means that 25% of the image is removed
///     on the given edge.
///     An image which is fully visible will have 0 for all values.
/// </remarks>
/// <param name="left">Portion of image along left edge of source picture which will not be displayed.</param>
/// <param name="right">Portion of image along right edge of source picture source picture which will not be displayed.</param>
/// <param name="top">Portion of image from top edge of source picture source picture which will not be displayed.</param>
/// <param name="bottom">Portion of image from bottom edge of source picture source picture which will not be displayed.</param>
public record CroppingFrame(decimal left, decimal right, decimal top, decimal bottom)
{
    /// <summary>
    ///     Set the cropping frame values onto the supplied source rectangle.
    /// </summary>
    /// <param name="aSrcRect">Rectange to be updated with our values.</param>
    public void UpdateSourceRectangle(A.SourceRectangle aSrcRect)
    {
        aSrcRect.Left = ToHundredThousandths(this.left);
        aSrcRect.Right = ToHundredThousandths(this.right);
        aSrcRect.Top = ToHundredThousandths(this.top);
        aSrcRect.Bottom = ToHundredThousandths(this.bottom);        
    }

    /// <summary>
    ///     Parse a string value into a cropping frame.
    /// </summary>
    /// <param name="input">All four frame values separated by commas.</param>
    /// <returns>Parsed frame.</returns>
    public static CroppingFrame Parse(string input)
    {        
        var split = input.Split(',');
        if (split.Length != 4)
        {
            throw new SCException("Must supply four numbers");
        }

        return new CroppingFrame(
            Decimal.Parse(split[0].Trim()),
            Decimal.Parse(split[1].Trim()),
            Decimal.Parse(split[2].Trim()),
            Decimal.Parse(split[3].Trim()));
    }

    /// <summary>
    ///     Convert a source rectangle to a cropping frame.
    /// </summary>
    /// <param name="aSrcRect">Source rectangle which contains the needed frame.</param>
    /// <returns>Resulting frame.</returns>
    public static CroppingFrame FromSourceRectangle(A.SourceRectangle? aSrcRect)
    {
        if (aSrcRect is null)
        {
            return new CroppingFrame(0,0,0,0);
        }

        return new CroppingFrame(
            FromHundredThousandths(aSrcRect.Left),
            FromHundredThousandths(aSrcRect.Right),
            FromHundredThousandths(aSrcRect.Top),
            FromHundredThousandths(aSrcRect.Bottom));
    }

    private static decimal FromHundredThousandths(Int32Value? int32) => int32 is not null ? int32 / 100000m : 0;

    private static Int32Value? ToHundredThousandths(decimal input) => 
        input == 0 ? null : Convert.ToInt32(input * 100000m);
}
