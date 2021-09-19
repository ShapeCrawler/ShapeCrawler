using ShapeCrawler.Shapes;

namespace ShapeCrawler.Audio
{
    public interface IAudioShape : IShape
    {
        /// <summary>
        ///     Gets audio's data in bytes.
        /// </summary>
        byte[] BinaryData { get; } // TODO: add setter

        // TODO: add ContentType property containing MIME type of audio
    }
}
