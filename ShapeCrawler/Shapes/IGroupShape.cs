using ShapeCrawler.Collections;
using ShapeCrawler.Shapes;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a group shape on a slide.
/// </summary>
public interface IGroupShape : IShape
{
    /// <summary>
    ///     Gets the collection of shapes inside the group.
    /// </summary>
    IGroupedShapeCollection Shapes { get; }
}