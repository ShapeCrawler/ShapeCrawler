using ShapeCrawler.AutoShapes;
using ShapeCrawler.Collections;
// ReSharper disable CheckNamespace

namespace ShapeCrawler
{
    public interface IParagraph
    {
        /// <summary>
        ///     Gets or sets the the plain text of a paragraph.
        /// </summary>
        string Text { get; set; }

        /// <summary>
        ///     Gets collection of paragraph portions. Returns <c>NULL</c> if paragraph is empty.
        /// </summary>
        IPortionCollection Portions { get; }

        /// <summary>
        ///     Gets paragraph bullet. Returns <c>NULL</c> if bullet does not exist.
        /// </summary>
        Bullet Bullet { get; }
    }
}