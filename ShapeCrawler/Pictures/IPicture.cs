using ShapeCrawler.Factories.Drawing;
using ShapeCrawler.Models;

namespace ShapeCrawler.Pictures
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