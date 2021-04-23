using System.IO;
using ShapeCrawler.Drawing;

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
        ///     Gets custom data.
        /// </summary>
        string CustomData { get; set; }

        /// <summary>
        ///     Determines whether slide is hidden.
        /// </summary>
        bool Hidden { get; }

        /// <summary>
        ///     Hides slide.
        /// </summary>
        void Hide();

        void SaveScheme(Stream stream);

        void SaveScheme(string filePath);

#if DEBUG
        void SaveImage(string filePath);
#endif
    }
}