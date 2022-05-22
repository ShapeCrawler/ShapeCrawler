using System.IO;

namespace ShapeCrawler.Extensions
{
    public static class StreamExtensions // TODO: make internal for Release
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
    }
}