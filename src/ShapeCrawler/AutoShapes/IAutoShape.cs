

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents interface of AutoShape.
/// </summary>
public interface IAutoShape : IShape
{
    /// <summary>
    ///     Gets shape outline.
    /// </summary>
    IShapeOutline Outline { get; }
 
    /// <summary>
    ///     Gets shape fill. Returns <see langword="null"/> if the shape can not be filled, for example, a line.
    /// </summary>
    IShapeFill Fill { get; }
    
    /// <summary>
    ///     Gets text frame. Returns <see langword="null"/> if the shape is not a text holder.
    /// </summary>
    ITextFrame TextFrame { get; }

    /// <summary>
    ///     Gets value indicating whether the Auto Shape is text holder.
    /// </summary>
    /// <returns></returns>
    bool IsTextHolder();
    
    /// <summary>
    ///     Gets the rotation of the shape in degrees.
    /// </summary>
    double Rotation { get; }
}