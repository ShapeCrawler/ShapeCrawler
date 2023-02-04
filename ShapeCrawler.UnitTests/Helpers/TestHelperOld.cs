using System.IO;
using System.Reflection;

namespace ShapeCrawler.UnitTests.Helpers;

public static class TestHelperOld
{
    static TestHelperOld()
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

    public static MemoryStream GetStream(string fileName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var stream = assembly.GetResourceStream(fileName);
        var mStream = new MemoryStream();
        stream!.CopyTo(mStream);

        return mStream;
    }

    public static readonly float HorizontalResolution;
        
    public static readonly float VerticalResolution;
}