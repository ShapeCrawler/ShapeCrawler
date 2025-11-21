using System;
using System.Collections;
using System.Collections.Generic;
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

internal sealed class AudioVideoShapeCollection(
    ISlideShapeCollection shapes,
    PresentationImageFiles presentationImageFiles,
    SlidePart slidePart
) : ISlideShapeCollection
{
    #region Shapes Properties

    public int Count => shapes.Count;

    public IShape this[int index] => shapes[index];

    #endregion Shapes Properties

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
        A.Extents extents = new() { Cx = 609600L, Cy = 609600L };

        var transform2D = new A.Transform2D(offset, extents);
        A.PresetGeometry presetGeometry = new(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle };

        var shapeProperties = new P.ShapeProperties(transform2D, presetGeometry);
        var nonVisualPictureProperties = new P.NonVisualPictureProperties(
            nonVisualDrawingProperties,
            nonVisualPictureDrawingProperties,
            applicationNonVisualDrawingProperties);
        var pPicture = new P.Picture(nonVisualPictureProperties, blipFill, shapeProperties);

        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pPicture);

        DocumentFormat.OpenXml.Office2010.PowerPoint.CreationId creationId = new() { Val = (UInt32Value)3972997422U };
        creationId.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");
    }

    #region Shapes Public Methods

    public void AddPicture(Stream imageStream) => shapes.AddPicture(imageStream);

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

    public void AddPieChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName,
        string chartName
    ) => shapes.AddPieChart(x, y, width, height, categoryValues, seriesName, chartName);

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

    public void AddClusteredBarChart(
        int x,
        int y,
        int width,
        int height,
        IList<string> categories,
        IList<Presentations.DraftChart.SeriesData> seriesData,
        string chartName
    ) => shapes.AddClusteredBarChart(x, y, width, height, categories, seriesData, chartName);

    public IShape AddSmartArt(
        int x,
        int y,
        int width,
        int height,
        SmartArtType smartArtType
    ) => shapes.AddSmartArt(x, y, width, height, smartArtType);

    public IShape Group(IShape[] groupingShapes) => shapes.Group(groupingShapes);

    public IShape AddDateAndTime() => shapes.AddDateAndTime();

    public IShape AddFooter() => shapes.AddFooter();

    public IShape AddSlideNumber() => shapes.AddSlideNumber();

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