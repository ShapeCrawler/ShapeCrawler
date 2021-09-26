using ShapeCrawler.Shapes;

namespace ShapeCrawler.Video
{
    public interface IVideoShape : IShape
    {
        /// <summary>
        ///     Gets video's data in bytes.
        /// </summary>
        byte[] BinaryData { get; }
    }
}
