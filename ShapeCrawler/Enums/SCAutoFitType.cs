namespace ShapeCrawler;

/// <summary>
///     Autofit type.
/// </summary>
public enum SCAutoFitType
{
    /// <summary>
    ///     Do not Autofit.
    /// </summary>
    None = 0,

    /// <summary>
    ///     Shrink text on overflow.
    /// </summary>
    Shrink = 1,

    /// <summary>
    ///     Resize shape to fit text.
    /// </summary>
    Resize = 2
}