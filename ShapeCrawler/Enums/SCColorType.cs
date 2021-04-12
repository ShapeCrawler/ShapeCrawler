using System.Diagnostics.CodeAnalysis;

namespace ShapeCrawler
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public enum SCColorType
    {
        NotDefined = 0,
        RGB = 1,
        RGBPercentage = 2,
        HSL = 3,

        /// <summary>
        ///     Color which is taken from theme scheme.
        /// </summary>
        Scheme = 4,

        System = 5,

        /// <summary>
        ///     Color which is bound to one of a predefined collection of colors.
        /// </summary>
        Preset = 6
    }
}