#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Fill type.
/// </summary>
public enum FillType
{
    /// <summary>
    ///     No fill.
    /// </summary>
    NoFill,

    /// <summary>
    ///     Solid fill.
    /// </summary>
    Solid,

    /// <summary>
    ///     Gradient fill.
    /// </summary>
    Gradient,

    /// <summary>
    ///     Picture fill.
    /// </summary>
    Picture,

    /// <summary>
    ///     Pattern fill.
    /// </summary>
    Pattern,

    /// <summary>
    ///     Slide background fill.
    /// </summary>
    SlideBackground
}