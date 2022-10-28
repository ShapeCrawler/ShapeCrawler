using System.Diagnostics.CodeAnalysis;
using System.IO;

namespace ShapeCrawler.Extensions
{
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1600", MessageId = "Elements should be documented", Justification = "Will be converted to internal")]
    public static class StreamExtensions
    {
        public static void WriteFile(this Stream destStream, string filePath)
        {
            destStream.SetLength(0);
            using var sourceStream = File.OpenRead(filePath);
            sourceStream.CopyTo(destStream);
        }

        public static void SaveToFile(this Stream sourceStream, string filePath)
        {
            sourceStream.Position = 0;
            using var destStream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
            sourceStream.CopyTo(destStream);
        }

        public static byte[] ToArray(this Stream stream)
        {
            if (stream is MemoryStream inputMs)
            {
                return inputMs.ToArray();
            }

            var ms = new MemoryStream();
            stream.CopyTo(ms);

            return ms.ToArray();
        }
    }
}