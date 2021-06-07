using ShapeCrawler.Shapes;
// ReSharper disable CheckNamespace

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a picture shape on a slide.
    /// </summary>
    public interface IPicture : IShape
    {
        /// <summary>
        ///     Gets image.
        /// </summary>
        SCImage Image { get; }
    }
}