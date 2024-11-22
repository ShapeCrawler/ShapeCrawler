#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

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