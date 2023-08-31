using System.IO;
using System.Reflection;

namespace ShapeCrawler.Shared;

internal sealed class Assets
{
    private readonly Assembly assembly;

    internal Assets(Assembly assembly)
    {
        this.assembly = assembly;
    }
    
    internal MemoryStream StreamOf(string file)
    {
        var stream = assembly.GetManifestResourceStream($"ShapeCrawler.Resources.{file}")!;
        var asset = new MemoryStream();
        stream.CopyTo(asset);
        asset.Position = 0;
        
        return asset;
    }

    internal string StringOf(string file)
    {
        return new StreamReader(this.StreamOf(file)).ReadToEnd();
    }
}