using System.IO;
using ShapeCrawler.Drawing;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a slide.
    /// </summary>
    public interface ISlide : IBaseSlide
    {
        /// <summary>
        ///     Gets slide number.
        /// </summary>
        int Number { get; }

        /// <summary>
        ///     Gets background image of the slide. Returns <c>NULL</c> if the slide does not have background.
        /// </summary>
        SCImage Background { get; }

        /// <summary>
        ///     Gets or sets custom data.
        /// </summary>
        string CustomData { get; set; }

        /// <summary>
        ///     Gets a value indicating whether slide hidden.
        /// </summary>
        bool Hidden { get; }

        ISlideLayout ParentSlideLayout { get; }

        /// <summary>
        ///     Hides slide.
        /// </summary>
        void Hide();

        /// <summary>
        ///     Saves slide scheme to stream.
        /// </summary>
        void SaveScheme(Stream stream);

        /// <summary>
        ///     Saves slide scheme to file.
        /// </summary>
        void SaveScheme(string filePath);

#if DEBUG
        void SaveImage(string filePath);
#endif
    }
}