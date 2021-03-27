using ShapeCrawler.Collections;
using ShapeCrawler.Drawing;

// ReSharper disable CheckNamespace

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a slide in a presentation.
    /// </summary>
    public interface ISlide
    {
        /// <summary>
        ///     Gets slide collection.
        /// </summary>
        ShapeCollection Shapes { get; }

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
    }
}