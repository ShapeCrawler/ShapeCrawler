using System.Collections.Generic;

namespace ShapeCrawler.Models
{
    /// <summary>
    ///     Represents a group shape on a slide.
    /// </summary>
    public interface IGroupShape : IShape
    {
        /// <summary>
        ///     Gets collection of the grouped.
        /// </summary>
        IReadOnlyCollection<IShape> Shapes { get; }
    }
}