using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ShapeCrawler.Tests.Helpers;

public static class AssemblyExtensions
{
    public static MemoryStream GetResourceStream(this Assembly assembly, string fileName)
    {
        var pattern = $@"\.{Regex.Escape(fileName)}";
        var path = assembly.GetManifestResourceNames().First(r =>
        {
            var matched = Regex.Match(r, pattern);
            return matched.Success;
        });
        var stream = assembly.GetManifestResourceStream(path);
        var mStream = new MemoryStream();
        stream!.CopyTo(mStream);
        mStream.Position = 0;

        return mStream;
    }
}