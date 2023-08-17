// ReSharper disable once CheckNamespace

using ShapeCrawler.Shapes;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a shape on a slide.
/// </summary>
public interface IShape : IShapeLocation
{
    /// <summary>
    ///     Gets or sets width of the shape in pixels.
    /// </summary>
    int Width { get; set; }

    /// <summary>
    ///     Gets or sets height of the shape in pixels.
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
    ///     Gets placeholder.
    /// </summary>
    IPlaceholder? Placeholder { get; }

    /// <summary>
    ///     Gets geometry form type of the shape.
    /// </summary>
    SCGeometry GeometryType { get; }

    /// <summary>
    ///     Gets or sets custom data string for the shape.
    /// </summary>
    string? CustomData { get; set; }

    /// <summary>
    ///     Gets shape type.
    /// </summary>
    SCShapeType ShapeType { get; }

    /// <summary>
    ///     Gets <see cref="IAutoShape"/>.    
    /// </summary>
    IAutoShape AsAutoShape();
}