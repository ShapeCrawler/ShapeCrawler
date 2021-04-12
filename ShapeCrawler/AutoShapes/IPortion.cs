// ReSharper disable CheckNamespace

using ShapeCrawler.AutoShapes;

namespace ShapeCrawler
{
    public interface IPortion
    {
        /// <summary>
        ///     Gets or sets paragraph portion text.
        /// </summary>
        string Text { get; set; }

        /// <summary>
        ///     Gets font.
        /// </summary>
        IFont Font { get; }

        /// <summary>
        ///     Removes portion from the paragraph.
        /// </summary>
        void Remove();
    }
}