// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a base class for user slide, layout or master slide.
    /// </summary>
    public interface IBaseSlide : IRemovable // TODO: what about using abstract class Slide instead; should be internal?
    {
        /// <summary>
        ///     Gets slide collection.
        /// </summary>
        IShapeCollection Shapes { get; }
    }
}