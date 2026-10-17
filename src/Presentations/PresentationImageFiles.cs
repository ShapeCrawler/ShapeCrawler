using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;

namespace ShapeCrawler.Presentations;

internal sealed class PresentationImageFiles(PresentationPart presentationPart)
{
    internal ImagePart? ImagePartByImageHashOrNull(string searchingImageHash)
    {
        foreach (var imagePart in this.ImageParts())
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

    private IEnumerable<ImagePart> ImageParts()
    {
        var masterSlideParts = presentationPart.SlideMasterParts;
        return presentationPart.SlideParts.SelectMany(slidePart => slidePart.ImageParts)
            .Concat(masterSlideParts.SelectMany(slideMasterPart => slideMasterPart.ImageParts))
            .Concat(masterSlideParts.SelectMany(slideMasterPart => slideMasterPart.SlideLayoutParts)
                .SelectMany(slideLayoutPart => slideLayoutPart.ImageParts));
    }
}