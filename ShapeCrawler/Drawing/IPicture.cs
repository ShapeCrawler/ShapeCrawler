using ShapeCrawler.Shapes;

namespace ShapeCrawler.Drawing
{
    /// <summary>
    ///     Represents a picture shape on a slide.
    /// </summary>
    public interface IPicture : IShape
    {
        /// <summary>
        ///     Gets image.
        /// </summary>
        ImageSc Image { get; }
    }
}