// ReSharper disable once CheckNamespace

using ShapeCrawler.Shapes;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a shape.
/// </summary>
public interface IShape : IPosition
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
    ///     Gets a value indicating whether shape is a placeholder.
    /// </summary>
    /// <returns></returns>
    bool IsPlaceholder { get; }
    
    /// <summary>
    ///     Gets placeholder.
    /// </summary>
    IPlaceholder Placeholder { get; }

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
    ///     Gets value indicating whether the shape has outline formatting.
    /// </summary>
    bool HasOutline { get; }
    
    /// <summary>
    ///     Gets shape outline.
    /// </summary>
    IShapeOutline Outline { get; }
 
    /// <summary>
    ///     Gets value indicating whether the shape has fill.
    /// </summary>
    bool HasFill { get; }
    
    /// <summary>
    ///     Gets shape fill. Returns <see langword="null"/> if the shape can not be filled, for example, a line.
    /// </summary>
    IShapeFill Fill { get; }
    
    /// <summary>
    ///     Gets value indicating whether the AutoShape is a text holder.
    /// </summary>
    /// <returns></returns>
    bool IsTextHolder { get; }
    
    /// <summary>
    ///     Gets text frame. Returns <see langword="null"/> if the shape is not a text holder.
    /// </summary>
    ITextFrame TextFrame { get; }
    
    /// <summary>
    ///     Gets the rotation of the shape in degrees.
    /// </summary>
    double Rotation { get; }

    /// <summary>
    ///     Gets the table if the shape is a table. Use <see cref="IShape.ShapeType"/> property to check if the shape is a table.
    /// </summary>
    /// <returns></returns>
    ITable AsTable();
    
    /// <summary>
    ///     Gets the media shape which is an audio or video.
    ///     Use <see cref="IShape.ShapeType"/> property to check if the shape is an <see cref="SCShapeType.Audio"/> or <see cref="SCShapeType.Video"/>.
    /// </summary>
    /// <returns></returns>
    IMediaShape AsMedia();
}