using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Presentations;

internal sealed class MediaCollection
{
    private readonly Dictionary<string, ImagePart> imagePartByHash = [];

    internal bool TryGetImagePart(string hash, out ImagePart part) => this.imagePartByHash.TryGetValue(hash, out part!);

    internal void SetImagePart(string hash, ImagePart part) => this.imagePartByHash[hash] = part;
}