using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ShapeCrawler.Extensions;

internal static class AssemblyExtensions
{
    internal static Stream GetStream(this Assembly assembly, string fileName)
    {
        var path = assembly.GetManifestResourceNames().First(r => r.EndsWith(fileName, StringComparison.Ordinal));
        var stream = assembly.GetManifestResourceStream(path);

        return stream!;
    }
}