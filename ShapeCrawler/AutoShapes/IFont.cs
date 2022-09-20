using ShapeCrawler.Drawing;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents font.
    /// </summary>
    public interface IFont
    {
        /// <summary>
        ///     Gets or sets font name.
        /// </summary>
        string Name { get; set; }

        /// <summary>
        ///     Gets or sets font size in points.
        /// </summary>
        int Size { get; set; }

        /// <summary>
        ///     Gets or sets a value indicating whether font's width is bold.
        /// </summary>
        bool IsBold { get; set; }

        /// <summary>
        ///     Gets or sets a value indicating whether font is italic.
        /// </summary>
        bool IsItalic { get; set; }

        /// <summary>
        ///     Gets or sets a value underline 
        /// </summary>
        DocumentFormat.OpenXml.Drawing.TextUnderlineValues Underline { get; set; }

        /// <summary>
        ///     Gets font color.
        /// </summary>
        IColorFormat ColorFormat { get; }

        /// <summary>
        ///     Gets value indicating whether font size can be changed.
        /// </summary>
        bool CanChangeSize();
    }
}