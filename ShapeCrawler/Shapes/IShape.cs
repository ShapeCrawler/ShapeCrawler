using ShapeCrawler.Placeholders;

namespace ShapeCrawler.Shapes;

/// <summary>
///     Represents a shape on a slide.
/// </summary>
public interface IShape
{
    /// <summary>
    ///     Gets or sets x-coordinate of the upper-left corner of the shape.
    /// </summary>
    int X { get; set; }

    /// <summary>
    ///     Gets or sets y-coordinate of the upper-left corner of the shape.
    /// </summary>
    int Y { get; set; }

    /// <summary>
    ///     Gets or sets width of the shape.
    /// </summary>
    int Width { get; set; }

    /// <summary>
    ///     Gets or sets height of the shape.
    /// </summary>
    int Height { get; set; }

    /// <summary>
    ///     Gets identifier of the shape.
    /// </summary>
    int Id { get; }

    /// <summary>
    ///     Gets name of the shape.
    /// </summary>
    string Name { get; }

    /// <summary>
    ///     Gets a value indicating whether shape is hidden.
    /// </summary>
    bool Hidden { get; }

    /// <summary>
    ///     Gets placeholder if shape is a placeholder, otherwise <see langword="null"/>.
    /// </summary>
    IPlaceholder? Placeholder { get; }

    /// <summary>
    ///     Gets geometry form type of the shape.
    /// </summary>
    SCGeometry GeometryType { get; }

    /// <summary>
    ///     Gets or sets custom data for the shape.
    /// </summary>
    string? CustomData { get; set; }

    /// <summary>
    ///     Gets shape type.
    /// </summary>
    SCShapeType ShapeType { get; }

    /// <summary>
    ///     Gets slide object.
    /// </summary>
    ISlideObject SlideObject { get; }
}