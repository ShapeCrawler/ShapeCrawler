using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ShapeCrawler.Tests.Helpers;

public static class TestHelper
{
    static TestHelper()
    {
        HorizontalResolution = 96;
        VerticalResolution = 96;
    }

    public static MemoryStream ToResizeableStream(this byte[] byteArray)
    {
        var stream = new MemoryStream();
        stream.Write(byteArray, 0, byteArray.Length);

        return stream;
    }

    public static MemoryStream GetStream(string file)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var path = assembly.GetManifestResourceNames().First(r => r.EndsWith(file, StringComparison.Ordinal));
        var stream = assembly.GetManifestResourceStream(path);
        var mStream = new MemoryStream();
        stream!.CopyTo(mStream);

        return mStream;
    }

    public static readonly float HorizontalResolution;
        
    public static readonly float VerticalResolution;
}