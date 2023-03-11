using System.IO;
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

    public static void SaveToTemp(IPresentation pres)
    {
        var folderPath = Path.Combine(@"c:\temp", TestContext.CurrentContext.Test.Name);
        var filePath = Path.Combine(folderPath, "result.pptx");
        Directory.CreateDirectory(folderPath);
        
        pres.SaveAs(filePath);
    }
}