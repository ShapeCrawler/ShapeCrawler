// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents base class for Slide, Slide Layout and Slide Master.
    /// </summary>
    internal abstract class SlideBase : IRemovable
    {
        /// <summary>
        ///     Gets slide collection.
        /// </summary>
        IShapeCollection Shapes { get; }

        public abstract bool IsRemoved { get; set; }

        public abstract void ThrowIfRemoved();
    }
}