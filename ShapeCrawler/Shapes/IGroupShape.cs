using System.Collections.Generic;

namespace ShapeCrawler.Shapes
{
    /// <summary>
    ///     Represents a group shape on a slide.
    /// </summary>
    public interface IGroupShape : IShape
    {
        /// <summary>
        ///     Gets the collection of shapes inside the group.
        /// </summary>
        IReadOnlyCollection<IShape> Shapes { get; }
    }
}