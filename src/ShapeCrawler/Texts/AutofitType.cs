#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Autofit type.
/// </summary>
public enum AutofitType
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