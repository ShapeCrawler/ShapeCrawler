using System.IO;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Drawing;

internal readonly ref struct WrappedImagePart
{
    private readonly ImagePart imagePart;

    internal WrappedImagePart(ImagePart imagePart)
    {
        this.imagePart = imagePart;
    }

    internal byte[] AsBytes()
    {
        var stream = this.imagePart.GetStream();

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