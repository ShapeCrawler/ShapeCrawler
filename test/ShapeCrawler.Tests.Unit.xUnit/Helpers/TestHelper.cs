using System.IO;
using System.Linq;
using System.Reflection;
using NUnit.Framework;

namespace ShapeCrawler.Tests.Unit.Helpers;

public static class TestHelper
{
    static TestHelper()
    {
        HorizontalResolution = 96;
        VerticalResolution = 96;
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