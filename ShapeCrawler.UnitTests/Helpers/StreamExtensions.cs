using System.IO;

namespace ShapeCrawler.UnitTests.Helpers;

public static class StreamExtensions
{
    public static void ToFile(this Stream stream, string filePath)
    {
        stream.Position = 0;
        using var destStream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        stream.CopyTo(destStream);
    }
}