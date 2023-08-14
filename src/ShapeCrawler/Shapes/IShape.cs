// ReSharper disable once CheckNamespace

using ShapeCrawler.Shapes;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

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
    ///     Gets or sets y-coordinate of the upper-left corner of the shape in pixels.
    /// </summary>
    int Y { get; set; }

    /// <summary>
    ///     Gets or sets width of the shape in pixels.
    /// </summary>
    int Width { get; set; }

    /// <summary>
    ///     Gets or sets height of the shape in pixels.
    /// </summary>
    int Height { get; set; }

    /// <summary>
    /// Gets the rotation of the shape in degrees.
    /// </summary>
    double Rotation { get; }

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
    ///     Gets parent Slide, SlideLayout or SlideMaster.
    /// </summary>
    ISlideStructure SlideStructure { get; }

    /// <summary>
    ///     Returns <see cref="IAutoShape"/> if shape is an auto shape, otherwise <see langword="null"/>.
    /// </summary>
    IAutoShape? AsAutoShape();
}