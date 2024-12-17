using System;
using ShapeCrawler.Exceptions;

namespace ShapeCrawler.Drawing;

/// <summary>
///     The portion within an image which is visible. 
/// </summary>
/// <remarks>
///     Values are a percentage of the whole image, e.g. `25` means that 25% of the image is removed
///     on the given edge.
///     An image which is fully visible will have 0 for all values.
/// </remarks>
/// <param name="Left">Percentage of image along left edge of source picture which will not be displayed.</param>
/// <param name="Right">Percentage of image along right edge of source picture source picture which will not be displayed.</param>
/// <param name="Top">Percentage of image from top edge of source picture source picture which will not be displayed.</param>
/// <param name="Bottom">Percentage of image from bottom edge of source picture source picture which will not be displayed.</param>
public readonly record struct CroppingFrame(decimal Left, decimal Right, decimal Top, decimal Bottom)
{
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
}
