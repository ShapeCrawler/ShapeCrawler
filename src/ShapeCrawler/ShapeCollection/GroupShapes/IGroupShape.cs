// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a group shape on a slide.
/// </summary>
public interface IGroupShape : IShape
{
    /// <summary>
    ///     Gets the collection of grouped shapes.
    /// </summary>
    IShapes Shapes { get; }
}