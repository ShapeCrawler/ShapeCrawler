using System.Collections.Generic;

namespace ShapeCrawler.Models
{
    public interface IGroupShape : IShape
    {
        /// <summary>
        /// Gets the collection of shapes inside the group.
        /// </summary>
        IReadOnlyCollection<IShape> Shapes { get; }
    }
}