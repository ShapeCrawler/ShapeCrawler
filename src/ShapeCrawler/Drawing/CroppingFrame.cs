using System;
using ShapeCrawler.Exceptions;

namespace ShapeCrawler.Drawing;

/// <summary>
///     The portion within an image which is visible. 
/// </summary>
/// <remarks>
///     Values are a fraction of the whole image, e.g. `0.25` means that 25% of the image is removed
///     on the given edge.
///     An image which is fully visible will have 0 for all values.
/// </remarks>
/// <param name="left">Portion of image along left edge of source picture which will not be displayed</param>
/// <param name="right">Portion of image along right edge of source picture source picture which will not be displayed</param>
/// <param name="top">Portion of image from top edge of source picture source picture which will not be displayed</param>
/// <param name="bottom">Portion of image from bottom edge of source picture source picture which will not be displayed</param>
public record CroppingFrame(decimal left, decimal right, decimal top, decimal bottom)
{
    /// <summary>
    ///     Parse a string value into a cropping frame 
    /// </summary>
    /// <param name="input">All four frame values separated by commas</param>
    /// <returns>Parsed frame</returns>
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
            Decimal.Parse(split[3].Trim())
        );
    }
}
