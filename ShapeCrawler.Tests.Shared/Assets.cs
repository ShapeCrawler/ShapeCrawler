using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ShapeCrawler.Tests.Shared;

public static class Assets
{
    public static string GetPath(string fileName)
    {
        var stream = GetStream(fileName);
        var path = Path.GetTempFileName();
        File.WriteAllBytes(path, stream.ToArray());

        return path;
    }
    
    public static MemoryStream GetStream(string fileName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var pattern = $@"\.{Regex.Escape(fileName)}";
        var path = assembly.GetManifestResourceNames().First(r =>
        {
            var matched = Regex.Match(r, pattern);
            return matched.Success;
        });
        var stream = assembly.GetManifestResourceStream(path);
        var mStream = new MemoryStream();
        stream.CopyTo(mStream);
        mStream.Position = 0;

        return mStream;
    }
}