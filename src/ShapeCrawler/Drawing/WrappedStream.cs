using System.IO;

namespace ShapeCrawler.Drawing;

internal readonly ref struct WrappedStream
{
    private readonly Stream stream;

    internal WrappedStream(Stream stream)
    {
        this.stream = stream;
    }

    internal byte[] AsBytes()
    {
        var mStream = new MemoryStream();
        var buffer = new byte[1024];
        int read;
        while ((read = stream.Read(buffer, 0, buffer.Length)) > 0)
        {
            mStream.Write(buffer, 0, read);
        }

        stream.Close();

        return mStream.ToArray();
    }
}