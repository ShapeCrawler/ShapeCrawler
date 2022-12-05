using ShapeCrawler.Shapes;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents interface of AutoShape.
/// </summary>
public interface IAutoShape : IShape
{
    /// <summary>
    ///     Gets shape fill.
    /// </summary>
    IShapeFill Fill { get; }
    
    /// <summary>
    ///     Gets shape outline.
    /// </summary>
    IShapeOutline Outline { get; }

    /// <summary>
    ///     Gets text frame if shape is text holder, otherwise <see langword="null"/>.
    /// </summary>
    ITextFrame? TextFrame { get; }
}