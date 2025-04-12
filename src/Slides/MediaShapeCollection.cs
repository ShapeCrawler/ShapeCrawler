using System;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ImageMagick;
using ImageMagick.Formats;
using ShapeCrawler.Assets;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Presentations;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using A16 = DocumentFormat.OpenXml.Office2016.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable UseObjectOrCollectionInitializer
namespace ShapeCrawler.Slides;

internal sealed class MediaShapeCollection(
    IShapeCollection shapes,
    SlidePart slidePart,
    MediaCollection mediaCollection)
{
    internal void AddAudio(int x, int y, Stream audio) => this.AddAudio(x, y, audio, AudioType.Mp3);

    internal void AddAudio(int x, int y, Stream audio, AudioType type)
    {
        string? contentType;
        string? extension;
        switch (type)
        {
            case AudioType.Mp3:
                contentType = "audio/mpeg";
                extension = ".mp3";
                break;
            case AudioType.Wave:
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
        transform2D.Extents!.Cx = 609600L;
        transform2D.Extents!.Cy = 609600L;

        var nonVisualPictureProps = pPicture.NonVisualPictureProperties!;
        var nonVisualDrawingProps = GetPNonVisualDrawingProperties(pPicture);

        var hyperlinkOnClick = new A.HyperlinkOnClick { Id = string.Empty, Action = "ppaction://media" };
        nonVisualDrawingProps.Append(hyperlinkOnClick);
        nonVisualPictureProps.Append(new P.NonVisualPictureDrawingProperties());

        var applicationNonVisualDrawingProps = nonVisualPictureProps.ApplicationNonVisualDrawingProperties!;
        applicationNonVisualDrawingProps.Append(audioFromFile);
        applicationNonVisualDrawingProps.Append(appNonVisualDrawingPropsExtensionList);
    }

    internal void AddPicture(Stream imageStream)
    {
        try
        {
            using var image = CreateMagickImage(imageStream);
            var originalFormat = image.Format;
            
            EnsureSupportedImageFormat(image);
            HandleSvgFormat(image, originalFormat);
            
            uint width = image.Width, height = image.Height;
            if (originalFormat == MagickFormat.Svg)
            {
                ResizeSvgImageIfNeeded(image, ref width, ref height);
            }

            var rasterStream = PrepareRasterStream(image);
            imageStream.Position = rasterStream.Position = 0;

            var pPicture = originalFormat == MagickFormat.Svg
                ? this.CreateSvgPPicture(rasterStream, imageStream, "Picture")
                : this.CreatePPicture(rasterStream, "Picture", GetMimeType(image.Format));

            SetPictureTransform(pPicture, width, height);
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

    private static MagickImage CreateMagickImage(Stream imageStream)
    {
        return new MagickImage(
            imageStream,
            new MagickReadSettings { BackgroundColor = MagickColors.Transparent });
    }

    private static void EnsureSupportedImageFormat(MagickImage image)
    {
        MagickFormat[] supportedImageFormats =
        [
            MagickFormat.Jpeg,
            MagickFormat.Png,
            MagickFormat.Gif,
            MagickFormat.Tif,
            MagickFormat.Tiff,
            MagickFormat.Svg
        ];

        if (!supportedImageFormats.Contains(image.Format))
        {
            image.Format = image.HasAlpha ? MagickFormat.Png : MagickFormat.Jpeg;
        }
    }

    private static void HandleSvgFormat(MagickImage image, MagickFormat originalFormat)
    {
        if (originalFormat == MagickFormat.Svg)
        {
            image.Format = MagickFormat.Png;
            image.Density = new Density(384, DensityUnit.PixelsPerInch); // in PowerPoint, the resolution of the rasterized version of SVG is set to 384 PPI
        }
    }

    private static void ResizeSvgImageIfNeeded(MagickImage image, ref uint width, ref uint height)
    {
        if (height > 500 || width > 500)
        {
            height = height > 500 ? 500 : height;
            width = width > 500 ? 500 : width;
            width = height == 500 ? (uint)(height * image.Width / (decimal)image.Height) : width;
            height = width == 500 ? (uint)(width * image.Height / (decimal)image.Width) : height;
            image.Resize(width, height);
        }
    }

    private static MemoryStream PrepareRasterStream(MagickImage image)
    {
        var rasterStream = new MemoryStream();
        image.Settings.SetDefines(new PngWriteDefines { ExcludeChunks = PngChunkFlags.date });
        image.Settings.SetDefine("png:exclude-chunk", "tIME");
        image.Write(rasterStream);
        return rasterStream;
    }

    private static void SetPictureTransform(P.Picture pPicture, uint width, uint height)
    {
        var transform2D = pPicture.ShapeProperties!.Transform2D!;
        transform2D.Offset!.X = transform2D.Offset!.Y = 952500;
        transform2D.Extents!.Cx = new Pixels(width).AsHorizontalEmus();
        transform2D.Extents!.Cy = new Pixels(height).AsVerticalEmus();
    }

    private static string GetMimeType(MagickFormat format)
    {
        var mime = MagickFormatInfo.Create(format)?.MimeType;

        return mime ?? throw new SCException("Unsupported image format.");
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
        if (shapes.Any())
        {
            return shapes.Select(shape => shape.Id).Prepend(0).Max() + 1;
        }

        return 1;
    }

    private bool TryGetImageRId(string hash, out string imgPartRId)
    {
        if (mediaCollection.TryGetImagePart(hash, out var imagePart))
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
            (imgPartRId, var imagePart) = slidePart.AddImagePart(image, mimeType);
            mediaCollection.SetImagePart(hash, imagePart);
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

    private P.Picture CreateSvgPPicture(Stream rasterStream, Stream svgStream, string shapeName)
    {
        // The SVG Blip contains the vector data
        var svgHash = new ImageStream(svgStream).Base64Hash;
        if (!this.TryGetImageRId(svgHash, out var svgPartRId))
        {
            (svgPartRId, var svgPart) = slidePart.AddImagePart(svgStream, "image/svg+xml");
            mediaCollection.SetImagePart(svgHash, svgPart);
        }

        // There is a possible optimization here. If we've previously in this session rasterized
        // this SVG, we could look up the rasterized image by reference to its vector image so
        // we wouldn't have to rasterize it every time.

        // The A.Blip contains a raster representation of the vector image
        var imgHash = new ImageStream(rasterStream).Base64Hash;
        if (!this.TryGetImageRId(imgHash, out var imgPartRId))
        {
            (imgPartRId, var imagePart) = slidePart.AddImagePart(rasterStream, "image/png");
            mediaCollection.SetImagePart(imgHash, imagePart);
        }

        var nonVisualPictureProperties = new P.NonVisualPictureProperties();
        var shapeId = (uint)this.GetNextShapeId();
        var nonVisualDrawingProperties = new P.NonVisualDrawingProperties
        {
            Id = shapeId, Name = $"{shapeName} {shapeId}"
        };
        var nonVisualPictureDrawingProperties = new P.NonVisualPictureDrawingProperties();
        var appNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties();

        var aNonVisualDrawingPropertiesExtensionList =
            new A.NonVisualDrawingPropertiesExtensionList();

        var aNonVisualDrawingPropertiesExtension =
            new A.NonVisualDrawingPropertiesExtension { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

        var a16CreationId = new A16.CreationId();

        // "http://schemas.microsoft.com/office/drawing/2014/main"
        var a16 = DocumentFormat.OpenXml.Linq.A16.a16;
        a16CreationId.AddNamespaceDeclaration(nameof(a16), a16.NamespaceName);

        a16CreationId.Id = "{2BEA8DB4-11C1-B7BA-06ED-DC504E2BBEBE}";

        aNonVisualDrawingPropertiesExtension.AppendChild(a16CreationId);

        aNonVisualDrawingPropertiesExtensionList.AppendChild(aNonVisualDrawingPropertiesExtension);

        nonVisualDrawingProperties.AppendChild(aNonVisualDrawingPropertiesExtensionList);
        nonVisualPictureProperties.AppendChild(nonVisualDrawingProperties);
        nonVisualPictureProperties.AppendChild(nonVisualPictureDrawingProperties);
        nonVisualPictureProperties.AppendChild(appNonVisualDrawingProperties);

        var blipFill = new P.BlipFill();
        var aBlip = new A.Blip { Embed = imgPartRId };
        var aBlipExtensionList = new A.BlipExtensionList();
        var aBlipExtension = new A.BlipExtension { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };
        var a14UseLocalDpi = new A14.UseLocalDpi();

        // "http://schemas.microsoft.com/office/drawing/2010/main"
        var a14 = DocumentFormat.OpenXml.Linq.A14.a14;

        a14UseLocalDpi.AddNamespaceDeclaration(nameof(a14), a14.NamespaceName);
        a14UseLocalDpi.Val = false;
        aBlipExtension.AppendChild(a14UseLocalDpi);
        aBlipExtensionList.AppendChild(aBlipExtension);
        aBlipExtension = new A.BlipExtension { Uri = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" };
        var svgBlip = new DocumentFormat.OpenXml.Office2019.Drawing.SVG.SVGBlip() { Embed = svgPartRId };

        // "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
        var asvg = DocumentFormat.OpenXml.Linq.ASVG.asvg;

        svgBlip.AddNamespaceDeclaration(nameof(asvg), asvg.NamespaceName);
        aBlipExtension.AppendChild(svgBlip);
        aBlipExtensionList.AppendChild(aBlipExtension);
        aBlip.AppendChild(aBlipExtensionList);
        blipFill.AppendChild(aBlip);
        var aStretch = new A.Stretch();
        var aFillRectangle = new A.FillRectangle();
        aStretch.AppendChild(aFillRectangle);

        blipFill.AppendChild(aStretch);

        var transform2D = new A.Transform2D(
            new A.Offset { X = 0, Y = 0 },
            new A.Extents { Cx = 0, Cy = 0 });

        var presetGeometry = new A.PresetGeometry { Preset = A.ShapeTypeValues.Rectangle };

        var aAdjustValueList = new A.AdjustValueList();

        presetGeometry.AppendChild(aAdjustValueList);

        var shapeProperties = new P.ShapeProperties();
        shapeProperties.AppendChild(transform2D);
        shapeProperties.AppendChild(presetGeometry);

        var pPicture = new P.Picture();
        pPicture.AppendChild(nonVisualPictureProperties);
        pPicture.AppendChild(blipFill);
        pPicture.AppendChild(shapeProperties);

        slidePart.Slide.CommonSlideData!.ShapeTree!.AppendChild(pPicture);

        return pPicture;
    }
}