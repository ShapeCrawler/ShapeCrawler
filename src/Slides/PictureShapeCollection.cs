using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ImageMagick;
using ShapeCrawler.Extensions;
using ShapeCrawler.Presentations;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable UseObjectOrCollectionInitializer
namespace ShapeCrawler.Slides;

internal sealed class PictureShapeCollection
{
    private readonly OpenXmlPart openXmlPart;
    private readonly PresentationImageFiles imageFiles;

    internal PictureShapeCollection(OpenXmlPart openXmlPart, PresentationImageFiles imageFiles)
    {
        this.openXmlPart = openXmlPart;
        this.imageFiles = imageFiles;
    }

    private P.ShapeTree ShapeTree => this.openXmlPart switch
    {
        SlidePart slidePart => slidePart.Slide!.CommonSlideData!.ShapeTree!,
        SlideLayoutPart slideLayoutPart => slideLayoutPart.SlideLayout!.CommonSlideData!.ShapeTree!,
        SlideMasterPart slideMasterPart => slideMasterPart.SlideMaster!.CommonSlideData!.ShapeTree!,
        _ => throw new SCException("Pictures can only be added to slides, layouts, or masters."),
    };

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
                    svgPartRId = this.AddImagePart(svgStream, "image/svg+xml");
                }

                var imgHash = imageContent.Hash;
                if (!this.TryGetImageRId(imgHash, out var imgPartRId))
                {
                    imgPartRId = this.AddImagePart(rasterStream, "image/png");
                }

                var xmlPicture = new XmlPicture(this.ShapeTree, (uint)this.GetNextShapeId(), "Picture");
                pPicture = xmlPicture.CreateSvgPPicture(imgPartRId, svgPartRId);
            }
            else
            {
                var imageForPart = imageContent.IsOriginalFormatPreserved
                    ? imageContent.GetOriginalStream()
                    : imageContent.GetRasterStream();

                var hash = imageContent.Hash;
                if (!this.TryGetImageRId(hash, out var imgPartRId))
                {
                    imgPartRId = this.AddImagePart(imageForPart, imageContent.MimeType);
                }

                var xmlPicture = new XmlPicture(this.ShapeTree, (uint)this.GetNextShapeId(), "Picture");
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

    private string AddImagePart(Stream imageStream, string contentType) => this.openXmlPart switch
    {
        SlidePart slidePart => slidePart.AddImagePart(imageStream, contentType),
        SlideLayoutPart slideLayoutPart => slideLayoutPart.AddImagePart(imageStream, contentType),
        SlideMasterPart slideMasterPart => slideMasterPart.AddImagePart(imageStream, contentType),
        _ => throw new SCException("Pictures can only be added to slides, layouts, or masters."),
    };

    private string CreateRelationshipToPart(ImagePart imagePart) => this.openXmlPart switch
    {
        SlidePart slidePart => slidePart.CreateRelationshipToPart(imagePart),
        SlideLayoutPart slideLayoutPart => slideLayoutPart.CreateRelationshipToPart(imagePart),
        SlideMasterPart slideMasterPart => slideMasterPart.CreateRelationshipToPart(imagePart),
        _ => throw new SCException("Pictures can only be added to slides, layouts, or masters."),
    };

    private int GetNextShapeId()
    {
        var shapeIds = this.ShapeTree
            .Descendants<P.NonVisualDrawingProperties>()
            .Select(p => p.Id?.Value ?? 0U)
            .ToArray();

        return shapeIds.Length > 0 ? (int)shapeIds.Max() + 1 : 1;
    }

    private IEnumerable<ImagePart> ImageParts() => this.openXmlPart switch
    {
        SlidePart slidePart => slidePart.ImageParts,
        SlideLayoutPart slideLayoutPart => slideLayoutPart.ImageParts,
        SlideMasterPart slideMasterPart => slideMasterPart.ImageParts,
        _ => throw new SCException("Pictures can only be added to slides, layouts, or masters."),
    };

    private string GetRelationshipId(ImagePart imagePart) => this.openXmlPart switch
    {
        SlidePart slidePart => slidePart.GetIdOfPart(imagePart),
        SlideLayoutPart slideLayoutPart => slideLayoutPart.GetIdOfPart(imagePart),
        SlideMasterPart slideMasterPart => slideMasterPart.GetIdOfPart(imagePart),
        _ => throw new SCException("Pictures can only be added to slides, layouts, or masters."),
    };

    private bool TryGetImageRId(string hash, out string imgPartRId)
    {
        var imagePart = imageFiles.ImagePartByImageHashOrNull(hash);
        if (imagePart is not null)
        {
            var found = this.ImageParts().Any(part => part.Uri == imagePart.Uri);
            imgPartRId = found ? this.GetRelationshipId(imagePart) : this.CreateRelationshipToPart(imagePart);
            return true;
        }

        imgPartRId = string.Empty;
        return false;
    }
}