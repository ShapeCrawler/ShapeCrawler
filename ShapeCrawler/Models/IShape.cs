using ShapeCrawler.Placeholders;

namespace ShapeCrawler.Models
{
    /// <summary>
    ///     Represents a shape on a slide.
    /// </summary>
    public interface IShape
    {
        /// <summary>
        ///     Gets x-coordinate of the upper-left corner of the shape.
        /// </summary>
        long X { get; set; }

        /// <summary>
        ///     Gets y-coordinate of the upper-left corner of the shape.
        /// </summary>
        long Y { get; set; }

        /// <summary>
        ///     Gets width of the shape.
        /// </summary>
        long Width { get; set; }

        /// <summary>
        ///     Gets height of the shape.
        /// </summary>
        long Height { get; set; }

        /// <summary>
        ///     Gets identifier of the shape.
        /// </summary>
        int Id { get; }

        /// <summary>
        ///     Gets name of the shape.
        /// </summary>
        string Name { get; }

        /// <summary>
        ///     Determines whether shape is hidden.
        /// </summary>
        bool Hidden { get; }

        /// <summary>
        ///     Gets placeholder. Returns null if the shape is not placeholder.
        /// </summary>
        Placeholder Placeholder { get; }

        /// <summary>
        ///     Gets geometry form type of the shape.
        /// </summary>
        GeometryType GeometryType { get; }

        /// <summary>
        ///     Gets or sets custom data for the shape.
        /// </summary>
        string CustomData { get; set; }
    }
}