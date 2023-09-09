// ReSharper disable InconsistentNaming
// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Color source.
/// </summary>
public enum SCColorType
{
    /// <summary>
    ///     Color is not defined.
    /// </summary>
    NotDefined = 0,

    /// <summary>
    ///     RGB value is absolute.
    /// </summary>
    RGB = 1,

    /// <summary>
    ///     RGB value is based on percentage.
    /// </summary>
    RGBPercentage = 2,

    /// <summary>
    ///     Color is defined in "Hue Saturation and Lightness" (HSL) way.
    /// </summary>
    HSL = 3,

    /// <summary>
    ///     Color from theme scheme.
    /// </summary>
    Theme = 4,

    /// <summary>
    ///     Operating system predefined color.
    /// </summary>
    Standard = 5,

    /// <summary>
    ///     Color which is bound to one of a predefined collection of colors.
    /// </summary>
    Preset = 6
}