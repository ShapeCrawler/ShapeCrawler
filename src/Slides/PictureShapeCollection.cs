using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ImageMagick;
using ShapeCrawler.Extensions;
using ShapeCrawler.Presentations;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable UseObjectOrCollectionInitializer
namespace ShapeCrawler.Slides;

internal sealed class PictureShapeCollection(SlidePart slidePart, PresentationImageFiles imageFiles)
{
    internal void AddPicture(Stream imageStream)
    {
        try
        {
            var imageContent = new Image(imageStream);

            P.Picture pPicture;
            if (imageContent.IsSvg)
            {
                var rasterStream = imageContent.GetRasterStream();
                var svgStream = imageContent.GetOriginalStream();

                var svgHash = imageContent.SvgHash;
                if (!this.TryGetImageRId(svgHash, out var svgPartRId))
                {
                    svgPartRId = slidePart.AddImagePart(svgStream, "image/svg+xml");
                }

                var imgHash = imageContent.Hash;
                if (!this.TryGetImageRId(imgHash, out var imgPartRId))
                {
                    imgPartRId = slidePart.AddImagePart(rasterStream, "image/png");
                }

                var xmlPicture = new XmlPicture(slidePart, (uint)this.GetNextShapeId(), "Picture");
                pPicture = xmlPicture.CreateSvgPPicture(imgPartRId, svgPartRId);
            }
            else
            {
                var imageForPart =

                    // Preserve original bytes for supported formats to ensure deterministic dedup across slides
                    imageContent.IsOriginalFormatPreserved ? imageContent.GetOriginalStream() :

                    // For formats that we convert (e.g., WebP/AVIF/BMP), write a deterministic raster representation
                    imageContent.GetRasterStream();

                var hash = imageContent.Hash;
                if (!this.TryGetImageRId(hash, out var imgPartRId))
                {
                    imgPartRId = slidePart.AddImagePart(imageForPart, imageContent.MimeType);
                }

                var xmlPicture = new XmlPicture(slidePart, (uint)this.GetNextShapeId(), "Picture");
                pPicture = xmlPicture.CreatePPicture(imgPartRId);
            }

            XmlPicture.SetTransform(pPicture, imageContent.Width, imageContent.Height);
        }
        catch (Exception ex) when (ex is MagickDelegateErrorException mex && mex.Message.Contains("ghostscript"))
        {
            throw new SCException(
                "The stream is an image format that requires GhostScript which is not installed on your system.", ex);
        }
        catch (MagickException)
        {
            throw new SCException(
                "The stream is not an image or a non-supported image format. Contact us for support: https://github.com/ShapeCrawler/ShapeCrawler/discussions");
        }
    }

    private int GetNextShapeId()
    {
        var shapeIds = slidePart.Slide
            .Descendants<P.NonVisualDrawingProperties>()
            .Select(p => p.Id?.Value ?? 0U)
            .ToArray();

        return shapeIds.Length > 0 ? (int)shapeIds.Max() + 1 : 1;
    }

    private bool TryGetImageRId(string hash, out string imgPartRId)
    {
        var imagePart = imageFiles.ImagePartByImageHashOrNull(hash);
        if (imagePart is not null)
        {
            // Image already exists in the presentation so far.
            // Do we have a reference to it on this slide?
            var found = slidePart.ImageParts.Where(x => x.Uri == imagePart.Uri);

            // Yes, we already have a relationship with this part on this slide
            // So use that relationship ID
            imgPartRId = found.Any() ? slidePart.GetIdOfPart(imagePart) :

                // No, so let's create a relationship to it
                slidePart.CreateRelationshipToPart(imagePart);

            return true;
        }

        // Sorry, you'll need to create a new image part
        imgPartRId = string.Empty;

        return false;
    }
}