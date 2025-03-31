using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ImageMagick;
using ImageMagick.Formats;
using ShapeCrawler.Assets;
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

internal sealed class SlideShapeCollection : ISlideShapeCollection
{
    private const long DefaultTableWidthEmu = 8128000L;

    private static readonly MagickFormat[] SupportedImageFormats =
    [
        MagickFormat.Jpeg,
        MagickFormat.Png,
        MagickFormat.Gif,
        MagickFormat.Tif,
        MagickFormat.Tiff,
        MagickFormat.Svg
    ];

    private static readonly MagickFormat[] VectorImageFormats = [MagickFormat.Svg];

    private readonly SlidePart slidePart;
    private readonly IShapeCollection shapes;
    private readonly MediaCollection mediaCollection;

    internal SlideShapeCollection(SlidePart slidePart, IShapeCollection shapes, MediaCollection mediaCollection)
    {
        this.slidePart = slidePart;
        this.shapes = shapes;
        this.mediaCollection = mediaCollection;
    }

    public int Count => this.shapes.Count;

    public IShape this[int index] => this.shapes[index];

    public void Add(IShape shape)
    {
        var pShapeTree = this.slidePart.Slide.CommonSlideData!.ShapeTree!;
        switch (shape)
        {
            case Picture picture:
                picture.CopyTo(pShapeTree);
                break;
            case Shape addingShape:
                addingShape.CopyTo(pShapeTree);
                break;
            default:
                throw new SCException("Unsupported shape type for adding.");
        }
    }

    public void AddAudio(int x, int y, Stream audio) => this.AddAudio(x, y, audio, AudioType.Mp3);

    public void AddAudio(int x, int y, Stream audio, AudioType type)
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
        var sdkPresentationDocument = (PresentationDocument)this.slidePart.OpenXmlPackage;
        var mediaDataPart = sdkPresentationDocument.CreateMediaDataPart(contentType, extension);
        audio.Position = 0;
        mediaDataPart.FeedData(audio);
        var imageStream = new AssetCollection(Assembly.GetExecutingAssembly()).StreamOf("audio image.png");

        var audioRef = this.slidePart.AddAudioReferenceRelationship(mediaDataPart);
        var mediaRef = this.slidePart.AddMediaReferenceRelationship(mediaDataPart);

        var audioFromFile = new A.AudioFromFile() { Link = audioRef.Id };

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

    public void AddPicture(Stream image)
    {
        image.Position = 0;
        try
        {
            using var imageMagick = new MagickImage(
                image,
                new MagickReadSettings { BackgroundColor = MagickColors.Transparent });
            var originalFormat = imageMagick.Format;
            if (!SupportedImageFormats.Contains(imageMagick.Format))
            {
                imageMagick.Format = imageMagick.HasAlpha ? MagickFormat.Png : MagickFormat.Jpeg;
            }

            if (VectorImageFormats.Contains(imageMagick.Format))
            {
                imageMagick.Format = MagickFormat.Png;
                imageMagick.Density =
                    new Density(384, DensityUnit.PixelsPerInch); // in PowerPoint, the resolution of the rasterized version of SVG is set to 384 PPI
            }

            var width = imageMagick.Width;
            var height = imageMagick.Height;

            if (height > 500)
            {
                height = 500;
                width = (uint)(height * imageMagick.Width / (decimal)imageMagick.Height);
            }

            if (width > 500)
            {
                width = 500;
                height = (uint)(width * imageMagick.Height / (decimal)imageMagick.Width);
            }

            if (width == 500 || height == 500)
            {
                imageMagick.Resize(width, height);
            }

            imageMagick.Settings.SetDefines(
                new PngWriteDefines { ExcludeChunks = PngChunkFlags.date });

            imageMagick.Settings.SetDefine("png:exclude-chunk", "tIME");

            var rasterStream = new MemoryStream();
            imageMagick.Write(rasterStream);
            image.Position = 0;
            rasterStream.Position = 0;
            var pPicture = VectorImageFormats.Contains(originalFormat)
                ? this.CreateSvgPPicture(rasterStream, image, "Picture")
                : this.CreatePPicture(rasterStream, "Picture", GetMimeType(imageMagick.Format));

            var widthEmu = new Pixels(width).AsHorizontalEmus();
            var heightEmu = new Pixels(height).AsVerticalEmus();
            var transform2D = pPicture.ShapeProperties!.Transform2D!;
            transform2D.Offset!.X = 952500;
            transform2D.Offset!.Y = 952500;
            transform2D.Extents!.Cx = widthEmu;
            transform2D.Extents!.Cy = heightEmu;
        }
        catch (MagickDelegateErrorException ex) when (ex.Message.Contains("ghostscript"))
        {
            throw new SCException(
                "The stream is an image format that requires GhostScript which is not installed on your system.", ex);
        }
        catch (MagickException)
        {
            throw new SCException(
                "The stream is not an image or a non-supported image format. You can raise a discussion at https://github.com/ShapeCrawler/ShapeCrawler/discussions to find out about the possibilities supporting it.");
        }
    }

    public void AddPieChart(
        int x, 
        int y, 
        int width, 
        int height, 
        Dictionary<string, double> categoryValues,
        string seriesName)
    {
        if (seriesName == null)
        {
            throw new ArgumentNullException(nameof(seriesName));
        }

        new SCSlidePart(this.slidePart).AddPieChart(x, y, width, height, categoryValues, seriesName);
    }

    public void AddVideo(int x, int y, Stream stream)
    {
        var presDocument = (PresentationDocument)this.slidePart.OpenXmlPackage;
        var xEmu = new Points(x).AsEmus();
        var yEmu = new Points(y).AsEmus();

        var mediaDataPart = presDocument.CreateMediaDataPart("video/mp4", ".mp4");

        stream.Position = 0;
        mediaDataPart.FeedData(stream);
        var imgPartRId = $"rId{Guid.NewGuid().ToString().Replace("-", string.Empty)[..5]}";
        var imagePart = this.slidePart.AddNewPart<ImagePart>("image/png", imgPartRId);
        var imageStream = new AssetCollection(Assembly.GetExecutingAssembly()).StreamOf("video image.bmp");
        imagePart.FeedData(imageStream);
        var videoRr = this.slidePart.AddVideoReferenceRelationship(mediaDataPart);
        var mediaRr = this.slidePart.AddMediaReferenceRelationship(mediaDataPart);

        var pPicture = new P.Picture();

        P.NonVisualPictureProperties nonVisualPictureProperties = new();

        var shapeId = (uint)this.shapes.Max(sp => sp.Id) + 1;
        P.NonVisualDrawingProperties nonVisualDrawingProperties = new() { Id = shapeId, Name = $"Video{shapeId}" };
        var hyperlinkOnClick = new A.HyperlinkOnClick { Id = string.Empty, Action = "ppaction://media" };

        A.NonVisualDrawingPropertiesExtensionList
            nonVisualDrawingPropertiesExtensionList = new();

        A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension =
            new() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

        nonVisualDrawingPropertiesExtensionList.Append(nonVisualDrawingPropertiesExtension);

        nonVisualDrawingProperties.Append(hyperlinkOnClick);
        nonVisualDrawingProperties.Append(nonVisualDrawingPropertiesExtensionList);

        P.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties = new();
        var pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

        nonVisualPictureDrawingProperties.Append(pictureLocks1);

        P.ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties = new();
        var videoFromFile = new A.VideoFromFile { Link = videoRr.Id };

        P.ApplicationNonVisualDrawingPropertiesExtensionList
            applicationNonVisualDrawingPropertiesExtensionList = new();

        P.ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension =
            new() { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };

        var media1 = new DocumentFormat.OpenXml.Office2010.PowerPoint.Media() { Embed = mediaRr.Id };
        media1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        applicationNonVisualDrawingPropertiesExtension.Append(media1);

        applicationNonVisualDrawingPropertiesExtensionList.Append(applicationNonVisualDrawingPropertiesExtension);

        applicationNonVisualDrawingProperties.Append(videoFromFile);
        applicationNonVisualDrawingProperties.Append(applicationNonVisualDrawingPropertiesExtensionList);

        nonVisualPictureProperties.Append(nonVisualDrawingProperties);
        nonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
        nonVisualPictureProperties.Append(applicationNonVisualDrawingProperties);

        P.BlipFill blipFill = new();
        A.Blip blip = new() { Embed = imgPartRId };
        A.Stretch stretch = new();
        A.FillRectangle fillRectangle = new();
        stretch.Append(fillRectangle);
        blipFill.Append(blip);
        blipFill.Append(stretch);

        P.ShapeProperties shapeProperties = new();

        A.Transform2D transform2D = new();
        A.Offset offset = new() { X = xEmu, Y = yEmu };
        A.Extents extents = new() { Cx = 609600L, Cy = 609600L };

        transform2D.Append(offset);
        transform2D.Append(extents);

        A.PresetGeometry presetGeometry = new() { Preset = A.ShapeTypeValues.Rectangle };
        A.AdjustValueList adjustValueList = new();

        presetGeometry.Append(adjustValueList);

        shapeProperties.Append(transform2D);
        shapeProperties.Append(presetGeometry);

        pPicture.Append(nonVisualPictureProperties);
        pPicture.Append(blipFill);
        pPicture.Append(shapeProperties);

        this.slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pPicture);

        DocumentFormat.OpenXml.Office2010.PowerPoint.CreationId creationId1 = new() { Val = (UInt32Value)3972997422U };
        creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");
    }

    public void AddShape(int x, int y, int width, int height, Geometry geometry = Geometry.Rectangle)
    {
        var xml = new AssetCollection(Assembly.GetExecutingAssembly()).StringOf("new rectangle.xml");
        var pShape = new P.Shape(xml);
        var nextShapeId = this.GetNextShapeId();
        this.slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pShape);

        var addedShape = this.shapes.Last<Shape>();
        addedShape.Name = geometry.ToString();
        addedShape.X = x;
        addedShape.Y = y;
        addedShape.Width = width;
        addedShape.Height = height;
        addedShape.Id = nextShapeId;
        addedShape.GeometryType = geometry;
    }

    public void AddLine(string xml)
    {
        var newPConnectionShape = new P.ConnectionShape(xml);

        this.slidePart.Slide.CommonSlideData!.ShapeTree!.Append(newPConnectionShape);
    }

    public void AddLine(int startPointX, int startPointY, int endPointX, int endPointY)
    {
        var xml = new AssetCollection(Assembly.GetExecutingAssembly()).StringOf("new line.xml");
        var pConnectionShape = new P.ConnectionShape(xml);
        this.slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pConnectionShape);

        var deltaY = endPointY - startPointY;
        var cx = endPointX;

        var cy = endPointY;
        if (deltaY == 0)
        {
            cy = 0;
        }

        if (startPointX == endPointX)
        {
            cx = 0;
        }

        var x = startPointX;
        var y = startPointY;
        var flipV = false;
        var flipH = false;
        if (startPointX > endPointX && endPointY > startPointY)
        {
            x = endPointX;
            y = startPointY;
            cx = startPointX - endPointX;
            cy = endPointY;
            flipH = true;
        }
        else if (startPointX > endPointX && startPointY == endPointY)
        {
            x = startPointX;
            cx = Math.Abs(startPointX - endPointX);
            cy = 0;
        }
        else if (startPointY > endPointY)
        {
            y = startPointY;
            cy = endPointY;
            flipV = true;
        }

        if (cx == 0)
        {
            flipV = true;
        }

        if (startPointX > endPointX)
        {
            flipH = true;
        }

        var idAndName = this.GenerateIdAndName();
        pConnectionShape.NonVisualConnectionShapeProperties!.NonVisualDrawingProperties!.Id = (uint)idAndName.Item1;

        var xEmu = new Points(x).AsEmus();
        var yEmu = new Points(y).AsEmus();
        var cxEmu = new Points(cx).AsEmus();
        var cyEmu = new Points(cy).AsEmus();
        var aXfrm = pConnectionShape.ShapeProperties!.Transform2D!;
        aXfrm.Offset!.X = xEmu;
        aXfrm.Offset!.Y = yEmu;
        aXfrm.Extents!.Cx = cxEmu;
        aXfrm.Extents!.Cy = cyEmu;
        aXfrm.HorizontalFlip = new BooleanValue(flipH);
        aXfrm.VerticalFlip = new BooleanValue(flipV);
    }

    public void AddTable(int x, int y, int columnsCount, int rowsCount) =>
        this.AddTable(x, y, columnsCount, rowsCount, TableStyle.MediumStyle2Accent1);

    public void AddTable(int x, int y, int columnsCount, int rowsCount, ITableStyle style)
    {
        var shapeName = this.GenerateNextTableName();
        var xEmu = UnitConverter.HorizontalPixelToEmu(x);
        var yEmu = UnitConverter.VerticalPixelToEmu(y);
        var tableHeightEmu = Constants.DefaultRowHeightEmu * rowsCount;

        var graphicFrame = new P.GraphicFrame();
        var nonVisualGraphicFrameProperties = new P.NonVisualGraphicFrameProperties();
        var nonVisualDrawingProperties = new P.NonVisualDrawingProperties
        {
            Id = (uint)this.GetNextShapeId(), Name = shapeName
        };
        var nonVisualGraphicFrameDrawingProperties = new P.NonVisualGraphicFrameDrawingProperties();
        var applicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties();
        nonVisualGraphicFrameProperties.Append(nonVisualDrawingProperties);
        nonVisualGraphicFrameProperties.Append(nonVisualGraphicFrameDrawingProperties);
        nonVisualGraphicFrameProperties.Append(applicationNonVisualDrawingProperties);

        var offset = new A.Offset { X = xEmu, Y = yEmu };
        var extents = new A.Extents { Cx = DefaultTableWidthEmu, Cy = tableHeightEmu };
        var pTransform = new P.Transform(offset, extents);

        var graphic = new A.Graphic();
        var graphicData = new A.GraphicData { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" };
        var aTable = new A.Table();

        var tableProperties = new A.TableProperties { FirstRow = true, BandRow = true };
        var tableStyleId = new A.TableStyleId { Text = ((TableStyle)style).Guid };
        tableProperties.Append(tableStyleId);

        var tableGrid = new A.TableGrid();
        var gridWidthEmu = DefaultTableWidthEmu / columnsCount;
        for (var i = 0; i < columnsCount; i++)
        {
            var gridColumn = new A.GridColumn { Width = gridWidthEmu };
            tableGrid.Append(gridColumn);
        }

        aTable.Append(tableProperties);
        aTable.Append(tableGrid);
        for (var i = 0; i < rowsCount; i++)
        {
            aTable.AddRow(columnsCount);
        }

        graphicData.Append(aTable);
        graphic.Append(graphicData);
        graphicFrame.Append(nonVisualGraphicFrameProperties);
        graphicFrame.Append(pTransform);
        graphicFrame.Append(graphic);

        this.slidePart.Slide.CommonSlideData!.ShapeTree!.Append(graphicFrame);
    }

    public void Remove(IShape shape)
    {
        var removingShape = this.shapes.FirstOrDefault(sp => sp.Id == shape.Id) ??
                            throw new SCException("Shape is not found.");
        removingShape.Remove();
    }

    public T GetById<T>(int id)
        where T : IShape => this.shapes.GetById<T>(id);

    public T? TryGetById<T>(int id)
        where T : IShape => this.shapes.TryGetById<T>(id);

    public T GetByName<T>(string name)
        where T : IShape => this.shapes.GetByName<T>(name);

    public T? TryGetByName<T>(string name)
        where T : IShape => this.shapes.TryGetByName<T>(name);

    public IShape GetByName(string name) => this.shapes.GetByName(name);

    public T Last<T>()
        where T : IShape => this.shapes.Last<T>();

    public IEnumerator<IShape> GetEnumerator() => this.shapes.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

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
        if (this.shapes.Any())
        {
            return this.shapes.Select(shape => shape.Id).Prepend(0).Max() + 1;
        }

        return 1;
    }

    private (int, string) GenerateIdAndName()
    {
        var id = this.GetNextShapeId();

        return (id, $"Shape {id}");
    }

    private string GenerateNextTableName()
    {
        var maxOrder = 0;
        foreach (var shape in this.shapes)
        {
            var matchOrder = Regex.Match(shape.Name, "(?!Table )\\d+", RegexOptions.None, TimeSpan.FromSeconds(100));
            if (!matchOrder.Success)
            {
                continue;
            }

            var order = int.Parse(matchOrder.Value);
            if (order > maxOrder)
            {
                maxOrder = order;
            }
        }

        return $"Table {maxOrder + 1}";
    }

    private bool TryGetImageRId(string hash, out string imgPartRId)
    {
        if (this.mediaCollection.TryGetImagePart(hash, out var imagePart))
        {
            // Image already exists in the presentation so far.
            // Do we have a reference to it on this slide?
            var found = this.slidePart.ImageParts.Where(x => x.Uri == imagePart.Uri);
            if (found.Any())
            {
                // Yes, we already have a relationship with this part on this slide
                // So use that relationship ID
                imgPartRId = this.slidePart.GetIdOfPart(imagePart);
            }
            else
            {
                // No, so let's create a relationship to it
                imgPartRId = this.slidePart.CreateRelationshipToPart(imagePart);
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
            (imgPartRId, var imagePart) = this.slidePart.AddImagePart(image, mimeType);
            this.mediaCollection.SetImagePart(hash, imagePart);
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

        this.slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pPicture);

        return pPicture;
    }

    private P.Picture CreateSvgPPicture(Stream rasterStream, Stream svgStream, string shapeName)
    {
        // The SVG Blip contains the vector data
        var svgHash = new ImageStream(svgStream).Base64Hash;
        if (!this.TryGetImageRId(svgHash, out var svgPartRId))
        {
            (svgPartRId, var svgPart) = this.slidePart.AddImagePart(svgStream, "image/svg+xml");
            this.mediaCollection.SetImagePart(svgHash, svgPart);
        }

        // There is a possible optimization here. If we've previously in this session rasterized
        // this SVG, we could look up the rasterized image by reference to its vector image so
        // we wouldn't have to rasterize it every time.

        // The A.Blip contains a raster representation of the vector image
        var imgHash = new ImageStream(rasterStream).Base64Hash;
        if (!this.TryGetImageRId(imgHash, out var imgPartRId))
        {
            (imgPartRId, var imagePart) = this.slidePart.AddImagePart(rasterStream, "image/png");
            this.mediaCollection.SetImagePart(imgHash, imagePart);
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

        var aBlip = new A.Blip() { Embed = imgPartRId };

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

        this.slidePart.Slide.CommonSlideData!.ShapeTree!.AppendChild(pPicture);

        return pPicture;
    }
}