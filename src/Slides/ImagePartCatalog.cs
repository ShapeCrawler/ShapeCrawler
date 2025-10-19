using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Slides;

/// <summary>
/// Tracks image parts keyed by content hash to avoid duplication across slides.
/// </summary>
internal sealed class ImagePartCatalog
{
    private readonly IDictionary<string, ImagePart> imageParts = new Dictionary<string, ImagePart>();

    /// <summary>
    /// Adds existing image parts to the catalog to establish baseline deduplication state.
    /// </summary>
    /// <param name="existingParts">Image parts to index by content hash.</param>
    internal void SeedFrom(IEnumerable<ImagePart> existingParts)
    {
        foreach (var imagePart in existingParts)
        {
            var hash = this.ComputeHash(imagePart);
            if (!this.imageParts.ContainsKey(hash))
            {
                this.imageParts.Add(hash, imagePart);
            }
        }
    }

    /// <summary>
    /// Replaces duplicate image parts in the specified slide part with catalogued instances.
    /// </summary>
    /// <param name="slidePart">Slide part to inspect for duplicate images.</param>
    internal void Deduplicate(SlidePart slidePart)
    {
        foreach (var imagePart in slidePart.ImageParts.ToList())
        {
            var hash = this.ComputeHash(imagePart);
            if (this.imageParts.TryGetValue(hash, out var existingPart) && !ReferenceEquals(existingPart, imagePart))
            {
                var relId = slidePart.GetIdOfPart(imagePart);
                slidePart.DeletePart(imagePart);
                slidePart.AddPart(existingPart, relId);
                continue;
            }

            this.imageParts[hash] = imagePart;
        }
    }

    private string ComputeHash(OpenXmlPart part)
    {
        using var sha512 = SHA512.Create();
        using var stream = part.GetStream();
        stream.Position = 0;
        var hash = sha512.ComputeHash(stream);
        stream.Position = 0;
        return Convert.ToBase64String(hash);
    }
}
