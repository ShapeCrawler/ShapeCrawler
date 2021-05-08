using ShapeCrawler.Collections;

// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a slide in a presentation.
    /// </summary>
    public interface IBaseSlide // TODO: what about using abstract class Slide instead; should be internal?
    {
        /// <summary>
        ///     Gets slide collection.
        /// </summary>
        IShapeCollection Shapes { get; }

        void ThrowIfRemoved();
    }
}