using System;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Assets;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Presentations;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable UseObjectOrCollectionInitializer
namespace ShapeCrawler.Slides;

internal sealed class AudioVideoShapeCollection(SlidePart slidePart, PresentationImageFiles presentationImageFiles)
{
    public void AddAudio(int x, int y, Stream audio) => this.AddAudio(x, y, audio, AudioType.MP3);

    public void AddAudio(int x, int y, Stream audio, AudioType type)
    {
        string? contentType;
        string? extension;
        switch (type)
        {
            case AudioType.MP3:
                contentType = "audio/mpeg";
                extension = ".mp3";
                break;
            case AudioType.WAVE:
                contentType = "audio/wav";
                extension = ".wav";
                break;
            default:
                throw new SCException("Unsupported audio type.");
        }

        var xEmu = new Points(x).AsEmus();
        var yEmu = new Points(y).AsEmus();
        var sdkPresentationDocument = (PresentationDocument)slidePart.OpenXmlPackage;
        var mediaDataPart = sdkPresentationDocument.CreateMediaDataPart(contentType, extension);
        audio.Position = 0;
        mediaDataPart.FeedData(audio);
        var imageStream = new AssetCollection(Assembly.GetExecutingAssembly()).StreamOf("audio image.png");

        var audioRef = slidePart.AddAudioReferenceRelationship(mediaDataPart);
        var mediaRef = slidePart.AddMediaReferenceRelationship(mediaDataPart);

        var audioFromFile = new A.AudioFromFile { Link = audioRef.Id };

        var appNonVisualDrawingPropsExtensionList = new P.ApplicationNonVisualDrawingPropertiesExtensionList();

        var appNonVisualDrawingPropsExtension = new P.ApplicationNonVisualDrawingPropertiesExtension
        {
            Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"
        };

        var media = new DocumentFormat.OpenXml.Office2010.PowerPoint.Media { Embed = mediaRef.Id };
        media.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");
        appNonVisualDrawingPropsExtension.Append(media);
        appNonVisualDrawingPropsExtensionList.Append(appNonVisualDrawingPropsExtension);

        var pPicture = this.CreatePPicture(imageStream, "Audio");

        var transform2D = pPicture.ShapeProperties!.Transform2D!;
        transform2D.Offset!.X = xEmu;
        transform2D.Offset!.Y = yEmu;
        transform2D.Extents!.Cx = 609_600L;
        transform2D.Extents!.Cy = 609_600L;

        var nonVisualPictureProps = pPicture.NonVisualPictureProperties!;
        var nonVisualDrawingProps = GetPNonVisualDrawingProperties(pPicture);

        var hyperlinkOnClick = new A.HyperlinkOnClick { Id = string.Empty, Action = "ppaction://media" };
        nonVisualDrawingProps.Append(hyperlinkOnClick);
        nonVisualPictureProps.Append(new P.NonVisualPictureDrawingProperties());

        var applicationNonVisualDrawingProps = nonVisualPictureProps.ApplicationNonVisualDrawingProperties!;
        applicationNonVisualDrawingProps.Append(audioFromFile);
        applicationNonVisualDrawingProps.Append(appNonVisualDrawingPropsExtensionList);
    }

    public void AddVideo(int x, int y, Stream stream)
    {
        var presDocument = (PresentationDocument)slidePart.OpenXmlPackage;
        var mediaDataPart = presDocument.CreateMediaDataPart("video/mp4", ".mp4");
        mediaDataPart.FeedData(stream);
        var imagePartRId = $"rId{Guid.NewGuid().ToString().Replace("-", string.Empty)[..5]}";
        var imagePart = slidePart.AddNewPart<ImagePart>("image/png", imagePartRId);
        var imageStream = new AssetCollection(Assembly.GetExecutingAssembly()).StreamOf("video image.png");
        imagePart.FeedData(imageStream);
        var videoRr = slidePart.AddVideoReferenceRelationship(mediaDataPart);
        var mediaRr = slidePart.AddMediaReferenceRelationship(mediaDataPart);

        var shapeId = (uint)this.GetNextShapeId();
        P.NonVisualDrawingProperties nonVisualDrawingProperties = new() { Id = shapeId, Name = $"Video{shapeId}" };
        var hyperlinkOnClick = new A.HyperlinkOnClick { Id = string.Empty, Action = "ppaction://media" };

        A.NonVisualDrawingPropertiesExtensionList
            nonVisualDrawingPropertiesExtensionList = new();

        A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension =
            new() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

        nonVisualDrawingPropertiesExtensionList.Append(nonVisualDrawingPropertiesExtension);

        nonVisualDrawingProperties.Append(hyperlinkOnClick);
        nonVisualDrawingProperties.Append(nonVisualDrawingPropertiesExtensionList);

        var nonVisualPictureDrawingProperties =
            new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true });

        var videoFromFile = new A.VideoFromFile { Link = videoRr.Id };

        P.ApplicationNonVisualDrawingPropertiesExtensionList
            applicationNonVisualDrawingPropertiesExtensionList = new();

        var media = new DocumentFormat.OpenXml.Office2010.PowerPoint.Media { Embed = mediaRr.Id };
        media.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");
        var applicationNonVisualDrawingPropertiesExtension =
            new P.ApplicationNonVisualDrawingPropertiesExtension(media)
            {
                Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"
            };
        applicationNonVisualDrawingPropertiesExtensionList.Append(applicationNonVisualDrawingPropertiesExtension);
        var applicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties(
            videoFromFile,
            applicationNonVisualDrawingPropertiesExtensionList);

        P.BlipFill blipFill = new();
        A.Blip blip = new() { Embed = imagePartRId };
        A.Stretch stretch = new();
        A.FillRectangle fillRectangle = new();
        stretch.Append(fillRectangle);
        blipFill.Append(blip);
        blipFill.Append(stretch);

        var xEmu = new Points(x).AsEmus();
        var yEmu = new Points(y).AsEmus();
        A.Offset offset = new() { X = xEmu, Y = yEmu };
        A.Extents extents = new() { Cx = 609_600L, Cy = 609_600L };

        var transform2D = new A.Transform2D(offset, extents);
        A.PresetGeometry presetGeometry = new(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle };

        var shapeProperties = new P.ShapeProperties(transform2D, presetGeometry);
        var nonVisualPictureProperties = new P.NonVisualPictureProperties(
            nonVisualDrawingProperties,
            nonVisualPictureDrawingProperties,
            applicationNonVisualDrawingProperties);
        var pPicture = new P.Picture(nonVisualPictureProperties, blipFill, shapeProperties);

        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pPicture);
    }

    private static P.NonVisualDrawingProperties GetPNonVisualDrawingProperties(OpenXmlCompositeElement compositeElement)
    {
        return compositeElement switch
        {
            P.GraphicFrame pGraphicFrame => pGraphicFrame.NonVisualGraphicFrameProperties!.NonVisualDrawingProperties!,
            P.Shape pShape => pShape.NonVisualShapeProperties!.NonVisualDrawingProperties!,
            P.Picture pPicture => pPicture.NonVisualPictureProperties!.NonVisualDrawingProperties!,
            P.GroupShape pGroupShape => pGroupShape.NonVisualGroupShapeProperties!.NonVisualDrawingProperties!,
            P.ConnectionShape pCxnSp => pCxnSp.NonVisualConnectionShapeProperties!.NonVisualDrawingProperties!,
            _ => throw new SCException()
        };
    }

    private int GetNextShapeId()
    {
        var shapeIds = slidePart.Slide
            .Descendants<P.NonVisualDrawingProperties>()
            .Select(p => p.Id?.Value ?? 0U)
            .ToList();

        return shapeIds.Count > 0 ? (int)shapeIds.Max() + 1 : 1;
    }

    private bool TryGetImageRId(string hash, out string imgPartRId)
    {
        var imagePart = presentationImageFiles.ImagePartByImageHashOrNull(hash);
        if (imagePart is not null)
        {
            // Image already exists in the presentation so far.
            // Do we have a reference to it on this slide?
            var found = slidePart.ImageParts.Where(x => x.Uri == imagePart.Uri);
            if (found.Any())
            {
                // Yes, we already have a relationship with this part on this slide
                // So use that relationship ID
                imgPartRId = slidePart.GetIdOfPart(imagePart);
            }
            else
            {
                // No, so let's create a relationship to it
                imgPartRId = slidePart.CreateRelationshipToPart(imagePart);
            }

            return true;
        }

        // Sorry, you'll need to create a new image part
        imgPartRId = string.Empty;
        return false;
    }

    private P.Picture CreatePPicture(Stream image, string shapeName, string mimeType = "image/png")
    {
        var imageStream = new ImageStream(image);
        var hash = imageStream.Base64Hash;

        if (!this.TryGetImageRId(hash, out var imgPartRId))
        {
            imgPartRId = slidePart.AddImagePart(image, mimeType);
        }

        var nonVisualPictureProperties = new P.NonVisualPictureProperties();
        var shapeId = (uint)this.GetNextShapeId();
        var nonVisualDrawingProperties = new P.NonVisualDrawingProperties
        {
            Id = shapeId, Name = $"{shapeName} {shapeId}"
        };
        var nonVisualPictureDrawingProperties = new P.NonVisualPictureDrawingProperties();
        var appNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties();

        nonVisualPictureProperties.Append(nonVisualDrawingProperties);
        nonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
        nonVisualPictureProperties.Append(appNonVisualDrawingProperties);

        var blipFill = new P.BlipFill();
        var blip = new A.Blip { Embed = imgPartRId };
        var stretch = new A.Stretch();
        blipFill.Append(blip);
        blipFill.Append(stretch);

        var transform2D = new A.Transform2D(
            new A.Offset { X = 0, Y = 0 },
            new A.Extents { Cx = 0, Cy = 0 });

        var presetGeometry = new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle };
        var shapeProperties = new P.ShapeProperties();
        shapeProperties.Append(transform2D);
        shapeProperties.Append(presetGeometry);

        var pPicture = new P.Picture();
        pPicture.Append(nonVisualPictureProperties);
        pPicture.Append(blipFill);
        pPicture.Append(shapeProperties);

        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pPicture);

        return pPicture;
    }
}