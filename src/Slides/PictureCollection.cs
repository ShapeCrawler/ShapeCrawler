using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ImageMagick;
using ImageMagick.Formats;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Presentations;
using ShapeCrawler.Shapes;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using A16 = DocumentFormat.OpenXml.Office2016.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable UseObjectOrCollectionInitializer
namespace ShapeCrawler.Slides;

internal sealed class PictureCollection(
    ISlideShapeCollection shapes,
    PresentationImageFiles presentationImageFiles,
    SlidePart slidePart
) : ISlideShapeCollection
{
    #region Shapes Properties

    public int Count => shapes.Count;

    public IShape this[int index] => shapes[index];

    #endregion Shapes Properties

    public void AddPicture(Stream imageStream)
    {
        try
        {
            if (imageStream.CanSeek)
            {
                imageStream.Position = 0;
            }

            using var image = CreateMagickImage(imageStream);
            var originalFormat = image.Format;

            EnsureSupportedImageFormat(image);
            HandleSvgFormat(image, originalFormat);

            uint width = image.Width, height = image.Height;
            if (originalFormat == MagickFormat.Svg)
            {
                ResizeSvgImageIfNeeded(image, ref width, ref height);
            }

            P.Picture pPicture;
            if (originalFormat == MagickFormat.Svg)
            {
                var rasterStream = PrepareRasterStream(image);
                imageStream.Position = rasterStream.Position = 0;
                pPicture = this.CreateSvgPPicture(rasterStream, imageStream, "Picture");
            }
            else if (originalFormat is MagickFormat.Gif or MagickFormat.Jpeg or MagickFormat.Png or MagickFormat.Tif or MagickFormat.Tiff)
            {
                // Preserve original bytes for supported formats to ensure deterministic dedup across slides
                imageStream.Position = 0;
                pPicture = this.CreatePPicture(imageStream, "Picture", GetMimeType(originalFormat));
            }
            else
            {
                // For formats we convert (e.g., WebP/AVIF/BMP), write a deterministic raster representation
                var rasterStream = PrepareRasterStream(image);
                imageStream.Position = rasterStream.Position = 0;
                pPicture = this.CreatePPicture(rasterStream, "Picture", GetMimeType(image.Format));
            }

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

    #region Shapes Public Methods

    public void AddAudio(int x, int y, Stream audio) => this.AddAudio(x, y, audio, AudioType.MP3);

    public void AddAudio(int x, int y, Stream audio, AudioType type) => shapes.AddAudio(x, y, audio, type);

    public void AddVideo(int x, int y, Stream stream) => shapes.AddVideo(x, y, stream);

    public void Add(IShape addingShape) => shapes.Add(addingShape);

    public void AddShape(
        int x,
        int y,
        int width,
        int height,
        Geometry geometry = Geometry.Rectangle
    ) => shapes.AddShape(x, y, width, height, geometry);

    public void AddShape(
        int x,
        int y,
        int width,
        int height,
        Geometry geometry,
        string text
    ) => shapes.AddShape(x, y, width, height, geometry, text);

    public void AddLine(string xml) => shapes.AddLine(xml);

    public void AddLine(
        int startPointX,
        int startPointY,
        int endPointX,
        int endPointY
    ) => shapes.AddLine(startPointX, startPointY, endPointX, endPointY);

    public void AddTable(
        int x,
        int y,
        int columnsCount,
        int rowsCount
    ) => shapes.AddTable(x, y, columnsCount, rowsCount);

    public void AddTable(
        int x,
        int y,
        int columnsCount,
        int rowsCount,
        ITableStyle style
    ) => shapes.AddTable(x, y, columnsCount, rowsCount, style);

    public void AddPieChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName
    ) => shapes.AddPieChart(x, y, width, height, categoryValues, seriesName);

    public void AddBarChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName
    ) => shapes.AddBarChart(x, y, width, height, categoryValues, seriesName);

    public void AddScatterChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<double, double> pointValues,
        string seriesName
    ) => shapes.AddScatterChart(x, y, width, height, pointValues, seriesName);

    public void AddStackedColumnChart(
        int x,
        int y,
        int width,
        int height,
        IDictionary<string, IList<double>> categoryValues,
        IList<string> seriesNames
    ) => shapes.AddStackedColumnChart(x, y, width, height, categoryValues, seriesNames);

    public IShape AddSmartArt(
        int x,
        int y,
        int width,
        int height,
        SmartArtType smartArtType
    ) => shapes.AddSmartArt(x, y, width, height, smartArtType);

    public IShape Group(IShape[] groupingShapes)
    {
        throw new NotImplementedException();
    }

    public IEnumerator<IShape> GetEnumerator() => shapes.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => shapes.GetEnumerator();

    public IShape GetById(int id) => shapes.GetById<IShape>(id);

    public T GetById<T>(int id)
        where T : IShape => shapes.GetById<T>(id);

    public IShape Shape(string name) => shapes.Shape<IShape>(name);

    public T Shape<T>(string name)
        where T : IShape => shapes.Shape<T>(name);

    public T Last<T>()
        where T : IShape => shapes.Last<T>();

    #endregion Shapes Public Methods

    private static MagickImage CreateMagickImage(Stream imageStream)
    {
        var format = IsIco(imageStream)
            ? MagickFormat.Ico
            : MagickFormat.Unknown;

        return new MagickImage(
            imageStream,
            new MagickReadSettings
            {
                Format = format,
                BackgroundColor = MagickColors.Transparent,
            });
    }

    private static bool IsIco(Stream stream)
    {
        if (stream.Length < 6)
        {
            return false;
        }
    
        var originalPosition = stream.Position;
        stream.Seek(0, SeekOrigin.Begin);
    
        try
        {
            var header = new byte[6];
            var bytesRead = stream.Read(header, 0, 6);
        
            if (bytesRead < 6)
            {
                return false;
            }

            // ICO file signature:
            // Bytes 0-1: Reserved (must be 0x00 0x00)
            // Bytes 2-3: Image type (must be 0x01 0x00 for ICO)
            // Bytes 4-5: Number of images (must be > 0)
            // https://docs.fileformat.com/image/ico/#header
            return header[0] == 0x00 && 
                   header[1] == 0x00 && 
                   header[2] == 0x01 && 
                   header[3] == 0x00 &&
                   (header[4] > 0 || header[5] > 0);
        }
        finally
        {
            stream.Seek(originalPosition, SeekOrigin.Begin);
        }
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
            image.Density =
                new Density(384, DensityUnit.PixelsPerInch); // in PowerPoint, the resolution of the rasterized version of SVG is set to 384 PPI
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
            (imgPartRId, _) = slidePart.AddImagePart(image, mimeType);
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
            (svgPartRId, _) = slidePart.AddImagePart(svgStream, "image/svg+xml");
        }

        // There is a possible optimization here. If we've previously in this session rasterized
        // this SVG, we could look up the rasterized image by reference to its vector image so
        // we wouldn't have to rasterize it every time.

        // The A.Blip contains a raster representation of the vector image
        var imgHash = new ImageStream(rasterStream).Base64Hash;
        if (!this.TryGetImageRId(imgHash, out var imgPartRId))
        {
            (imgPartRId, _) = slidePart.AddImagePart(rasterStream, "image/png");
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