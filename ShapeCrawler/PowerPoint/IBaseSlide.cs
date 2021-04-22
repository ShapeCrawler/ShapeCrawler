using ShapeCrawler.Collections;
using ShapeCrawler.Drawing;

// ReSharper disable CheckNamespace

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a slide in a presentation.
    /// </summary>
    public interface IBaseSlide
    {
        /// <summary>
        ///     Gets slide collection.
        /// </summary>
        ShapeCollection Shapes { get; }

        /// <summary>
        ///     Gets slide number.
        /// </summary>
        int Number { get; } // TODO: Is it need for Slide Layout and Master

        /// <summary>
        ///     Gets background image of the slide. Returns <c>NULL</c> if the slide does not have background.
        /// </summary>
        SCImage Background { get; } // TODO: Is it need for Slide Layout and Master

        /// <summary>
        ///     Gets custom data.
        /// </summary>
        string CustomData { get; set; } // TODO: Is it need for Slide Layout and Master

        /// <summary>
        ///     Determines whether slide is hidden.
        /// </summary>
        bool Hidden { get; } // TODO: Is it need for Slide Layout and Master

        /// <summary>
        ///     Hides slide.
        /// </summary>
        void Hide(); // TODO: Is it need for Slide Layout and Master
    }

    internal interface IRemovable
    {
        bool IsRemoved { get; set; }
    }
}