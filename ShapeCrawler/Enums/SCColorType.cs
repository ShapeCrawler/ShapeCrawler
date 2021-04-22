using System.Diagnostics.CodeAnalysis;

namespace ShapeCrawler
{
    [SuppressMessage("ReSharper", "InconsistentNaming", Justification = "SC - ShapeCrawler")]
    public enum SCColorType
    {
        /// <summary>
        ///     Color is not defined.
        /// </summary>
        NotDefined = 0,

        /// <summary>
        ///     RGB color.
        /// </summary>
        RGB = 1,
        RGBPercentage = 2,
        HSL = 3,

        /// <summary>
        ///     Color from theme scheme.
        /// </summary>
        Scheme = 4,

        /// <summary>
        ///     System color.
        /// </summary>
        System = 5,

        /// <summary>
        ///     Color which is bound to one of a predefined collection of colors.
        /// </summary>
        Preset = 6
    }
}