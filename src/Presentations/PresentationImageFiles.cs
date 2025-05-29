using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;

namespace ShapeCrawler.Presentations;

internal sealed class PresentationImageFiles(IEnumerable<SlidePart> slideParts)
{
    internal ImagePart? ImagePartByImageHashOrNull(string searchingImageHash)
    {
        var imageParts = slideParts.SelectMany(slidePart => slidePart.ImageParts);
        foreach (var imagePart in imageParts)
        {
            using var stream = imagePart.GetStream();
            var hash = new ImageStream(stream).Base64Hash;
            if (hash == searchingImageHash)
            {
                return imagePart;
            }
        }

        return null;
    }
}