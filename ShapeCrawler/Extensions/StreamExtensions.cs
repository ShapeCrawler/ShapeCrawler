using System.IO;

namespace ShapeCrawler.Extensions
{
    public static class StreamExtensions
    {
        public static void WriteFile(this Stream destStream, string filePath)
        {
            using var sourceStream = File.OpenRead(filePath);
            sourceStream.CopyTo(destStream);
        }

        public static void SaveToFile(this Stream sourceStream, string filePath)
        {
            using var destStream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
            sourceStream.CopyTo(destStream);
        }
    }
}