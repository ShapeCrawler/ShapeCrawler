using ShapeCrawler.Drawing;

namespace ShapeCrawler.AutoShapes
{
    /// <summary>
    ///     Represents the font interface.
    /// </summary>
    public interface IFont
    {
        /// <summary>
        ///     Gets or sets font name.
        /// </summary>
        string Name { get; set; }

        /// <summary>
        ///     Gets or sets font size in EMUs.
        /// </summary>
        int Size { get; set; } // TODO: create test to verify font size of table cell's text portion

        /// <summary>
        ///     Gets or sets a value indicating whether font width is bold.
        /// </summary>
        bool IsBold { get; set; }

        /// <summary>
        ///     Gets or sets a value indicating whether font is italic.
        /// </summary>
        bool IsItalic { get; set; }

        /// <summary>
        ///     Gets color.
        /// </summary>
        IColorFormat ColorFormat { get; }

        /// <summary>
        ///     Gets value indicating whether font size can be changed.
        /// </summary>
        bool SizeCanBeChanged();
    }
}