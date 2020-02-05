using System.IO;

namespace SlideXML.Extensions
{
    /// <summary>
    /// Contains extension methods for <see cref="Stream"/> instance.
    /// </summary>
    public static class StreamExtensions
    {
        /// <summary>
        /// Seeks stream to begin if it is seekable.
        /// </summary>
        /// <param name="stream"></param>
        public static void SeekBegin(this Stream stream)
        {
            if (stream.CanSeek)
            {
                stream.Seek(0, SeekOrigin.Begin);
            }
        }
    }
}
