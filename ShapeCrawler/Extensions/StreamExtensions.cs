using System.Diagnostics.CodeAnalysis;
using System.IO;

namespace ShapeCrawler.Extensions;

[SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1600", MessageId = "Elements should be documented", Justification = "Will be converted to internal")]
internal static class StreamExtensions
{
    internal static byte[] ToArray(this Stream stream)
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