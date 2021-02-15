namespace ShapeCrawler.Models
{
    public interface IShape
    {
        /// <summary>
        /// Returns the x-coordinate of the upper-left corner of the shape.
        /// </summary>
        long X { get; set; }

        /// <summary>
        /// Returns the y-coordinate of the upper-left corner of the shape.
        /// </summary>
        long Y { get; set; }

        /// <summary>
        /// Returns the width of the shape.
        /// </summary>
        long Width { get; set; }

        /// <summary>
        /// Returns the height of the shape.
        /// </summary>
        long Height { get; set; }

        /// <summary>
        /// Returns an element identifier.
        /// </summary>
        int Id { get; }

        /// <summary>
        /// Gets an element name.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Determines whether the shape is hidden.
        /// </summary>
        bool Hidden { get; }

        Placeholder Placeholder { get; }

        GeometryType GeometryType { get; }

        string CustomData { get; set; }
    }
}