using System.Drawing;

namespace ShapeCrawler.Drawing
{
    /// <summary>
    ///     Represents the color interface.
    /// </summary>
    public interface IColorFormat
    {
        /// <summary>
        ///     Gets color type.
        /// </summary>
        SCColorType ColorType { get; }

#if DEBUG
        /// <summary>
        ///     Gets or sets ARGB color.
        /// </summary>
        Color Color { get; set; }
#else
        /// <summary>
        ///     Gets ARGB color.
        /// </summary>
        Color Color { get; }
#endif
    }
}