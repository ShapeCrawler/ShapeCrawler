using ShapeCrawler.AutoShapes;
using ShapeCrawler.Collections;

// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a paragraph.
    /// </summary>
    public interface IParagraph
    {
        /// <summary>
        /// Add text to paragraph.
        /// </summary>
        void AddText(string text);

        /// <summary>
        ///     Gets or sets paragraph text.
        /// </summary>
        string Text { get; set; }

        /// <summary>
        ///     Gets collection of paragraph portions.
        /// </summary>
        IPortionCollection Portions { get; }

        /// <summary>
        ///     Gets paragraph bullet. Returns <c>NULL</c> if bullet does not exist.
        /// </summary>
        Bullet Bullet { get; }

        /// <summary>
        ///     Gets or sets the text alignment.
        /// </summary>
        TextAlignment Alignment { get; set; }
    }
}