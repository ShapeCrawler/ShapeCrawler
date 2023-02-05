using System.IO;

namespace ShapeCrawler.Extensions;

internal static class ByteArrayExtensions
{
    internal static MemoryStream ToExpandableStream(this byte[] bytes)
    {
        var mStream = new MemoryStream();
        mStream.Write(bytes, 0, bytes.Length);

        return mStream;
    }
}