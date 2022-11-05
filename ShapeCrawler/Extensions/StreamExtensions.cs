using System.Diagnostics.CodeAnalysis;
using System.IO;

namespace ShapeCrawler.Extensions;

[SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1600", MessageId = "Elements should be documented", Justification = "Will be converted to internal")]
public static class StreamExtensions
{
    public static void ToFile(this Stream stream, string filePath)
    {
        stream.Position = 0;
        using var destStream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        stream.CopyTo(destStream);
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