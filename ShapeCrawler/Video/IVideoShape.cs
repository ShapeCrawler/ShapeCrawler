using ShapeCrawler.Shapes;

namespace ShapeCrawler.Video
{
    /// <summary>
    ///     Represents a shape containing video content.
    /// </summary>
    public interface IVideoShape : IShape
    {
        /// <summary>
        ///     Gets video's data in bytes.
        /// </summary>
        byte[] BinaryData { get; }
    }
}
