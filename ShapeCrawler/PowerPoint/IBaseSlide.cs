// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a base class for user slide, layout or master slide.
    /// </summary>
    public interface IBaseSlide : IRemovable
    {
        /// <summary>
        ///     Gets slide collection.
        /// </summary>
        IShapeCollection Shapes { get; }
    }
}