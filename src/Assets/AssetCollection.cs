using System.IO;
using System.Reflection;

namespace ShapeCrawler.Assets;

internal readonly ref struct AssetCollection(Assembly assembly)
{
    internal MemoryStream StreamOf(string file)
    {
        var stream = assembly.GetManifestResourceStream($"ShapeCrawler.Assets.{file}") !;
        var asset = new MemoryStream();
        stream.CopyTo(asset);
        asset.Position = 0;

        return asset;
    }

    internal string StringOf(string file) => new StreamReader(this.StreamOf(file)).ReadToEnd();
}
