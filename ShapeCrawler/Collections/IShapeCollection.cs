using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OneOf;
using ShapeCrawler.Collections;
using ShapeCrawler.Constants;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Media;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.SlideMasters;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a shape collection.
/// </summary>
public interface IShapeCollection : IEnumerable<IShape>
{
    /// <summary>
    ///     Gets the number of series items in the collection.
    /// </summary>
    int Count { get; }

    /// <summary>
    ///     Gets shape at the specified index.
    /// </summary>
    IShape this[int index] { get; }

    /// <summary>
    ///     Gets shape by identifier. Returns <see langword="null"/> if shape is not found.
    /// </summary>
    /// <typeparam name="T">The type of shape.</typeparam>
    T? GetById<T>(int shapeId)
        where T : IShape;

    /// <summary>
    ///     Gets shape by name. Returns <see langword="null"/> if shape is not found.
    /// </summary>
    /// <typeparam name="T">The type of shape.</typeparam>
    T? GetByName<T>(string shapeName)
        where T : IShape;

    /// <summary>
    ///     Gets shape by name. Returns <see langword="null"/> if shape is not found.
    /// </summary>
    IShape? GetByName(string shapeName);

    /// <summary>
    ///     Create a new audio shape from stream and adds it to the end of the collection.
    /// </summary>
    /// <param name="xPixel">The X coordinate for the left side of the shape.</param>
    /// <param name="yPixels">The Y coordinate for the left side of the shape.</param>
    /// <param name="mp3Stream">Audio stream data.</param>
    IAudioShape AddAudio(int xPixel, int yPixels, Stream mp3Stream);

    /// <summary>
    ///     Create a new video shape from stream and adds it to the end of the collection.
    /// </summary>
    /// <param name="x">X coordinate in pixels.</param>
    /// <param name="y">Y coordinate in pixels.</param>
    /// <param name="stream">Video stream data.</param>
    IVideoShape AddVideo(int x, int y, Stream stream);

    /// <summary>
    ///     Creates a new AutoShape.
    /// </summary>
    /// <param name="type">AutoShape type.</param>
    /// <param name="x">X coordinate in pixels.</param>
    /// <param name="y">Y coordinate in pixels.</param>
    /// <param name="width">Width in pixels.</param>
    /// <param name="height">Height in pixels.</param>
    IAutoShape AddAutoShape(SCAutoShapeType type, int x, int y, int width, int height);

    /// <summary>
    ///     Creates a new Table.
    /// </summary>
    ITable AddTable(int x, int y, int columns, int rows);

    /// <summary>
    ///     Removes specified shape.
    /// </summary>
    void Remove(IShape shape);
}

internal class ShapeCollection : LibraryCollection<IShape>, IShapeCollection
{
    private const long DefaultTableWidthEmu = 8128000L;
    private readonly P.ShapeTree shapeTree;
    private readonly OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject;

    private ShapeCollection(
        List<IShape> shapes, 
        P.ShapeTree shapeTree, 
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> slideObject)
        : base(shapes)
    {
        this.slideObject = slideObject;
        this.shapeTree = shapeTree;
    }

    public IAudioShape AddAudio(int xPixels, int yPixels, Stream mp3Stream)
    {
        long xEmu = UnitConverter.HorizontalPixelToEmu(xPixels);
        long yEmu = UnitConverter.VerticalPixelToEmu(yPixels);

        var slideBase = this.slideObject.Match(slide => slide as SlideObject, layout => layout, master => master);
        var mediaDataPart =
            slideBase.PresentationInternal.SDKPresentationInternal.CreateMediaDataPart("audio/mpeg", ".mp3");

        mp3Stream.Position = 0;
        mediaDataPart.FeedData(mp3Stream);
        string imgPartRId = $"rId{Guid.NewGuid().ToString().Replace("-", string.Empty).Substring(0, 5)}";
        var slidePart = (SlidePart)slideBase.TypedOpenXmlPart;
        var imagePart = slidePart!.AddNewPart<ImagePart>("image/png", imgPartRId);
        var imgStream = Assembly.GetExecutingAssembly().GetStream("audio-image.png");
        imgStream.Position = 0;
        imagePart.FeedData(imgStream);

        var audioRr = slidePart.AddAudioReferenceRelationship(mediaDataPart);
        var mediaRr = slidePart.AddMediaReferenceRelationship(mediaDataPart);

        P.Picture picture1 = new();

        P.NonVisualPictureProperties nonVisualPictureProperties1 = new();

        uint shapeId = (uint)this.CollectionItems.Max(sp => sp.Id) + 1;
        P.NonVisualDrawingProperties nonVisualDrawingProperties2 = new() { Id = shapeId, Name = $"Audio{shapeId}" };
        A.HyperlinkOnClick hyperlinkOnClick1 = new A.HyperlinkOnClick() { Id = "", Action = "ppaction://media" };

        A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new();

        A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 =
            new() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

        OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement(
            "<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{2FF36D28-5328-4DA3-BF85-A2B65D7EE127}\" />");

        nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

        nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

        nonVisualDrawingProperties2.Append(hyperlinkOnClick1);
        nonVisualDrawingProperties2.Append(nonVisualDrawingPropertiesExtensionList1);

        P.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new();
        A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

        nonVisualPictureDrawingProperties1.Append(pictureLocks1);

        P.ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new();
        A.AudioFromFile audioFromFile1 = new A.AudioFromFile() { Link = audioRr.Id };

        P.ApplicationNonVisualDrawingPropertiesExtensionList
            applicationNonVisualDrawingPropertiesExtensionList1 = new();

        P.ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 =
            new() { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };

        P14.Media media1 = new P14.Media() { Embed = mediaRr.Id };
        media1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        applicationNonVisualDrawingPropertiesExtension1.Append(media1);

        applicationNonVisualDrawingPropertiesExtensionList1.Append(applicationNonVisualDrawingPropertiesExtension1);

        applicationNonVisualDrawingProperties2.Append(audioFromFile1);
        applicationNonVisualDrawingProperties2.Append(applicationNonVisualDrawingPropertiesExtensionList1);

        nonVisualPictureProperties1.Append(nonVisualDrawingProperties2);
        nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
        nonVisualPictureProperties1.Append(applicationNonVisualDrawingProperties2);

        P.BlipFill blipFill1 = new();
        A.Blip blip1 = new() { Embed = imgPartRId };

        A.Stretch stretch1 = new();
        A.FillRectangle fillRectangle1 = new();

        stretch1.Append(fillRectangle1);

        blipFill1.Append(blip1);
        blipFill1.Append(stretch1);

        P.ShapeProperties shapeProperties1 = new();

        A.Transform2D transform2D1 = new();
        A.Offset offset2 = new() { X = xEmu, Y = yEmu };
        A.Extents extents2 = new() { Cx = 609600L, Cy = 609600L };

        transform2D1.Append(offset2);
        transform2D1.Append(extents2);

        A.PresetGeometry presetGeometry1 = new() { Preset = A.ShapeTypeValues.Rectangle };
        A.AdjustValueList adjustValueList1 = new();

        presetGeometry1.Append(adjustValueList1);

        shapeProperties1.Append(transform2D1);
        shapeProperties1.Append(presetGeometry1);

        picture1.Append(nonVisualPictureProperties1);
        picture1.Append(blipFill1);
        picture1.Append(shapeProperties1);

        this.shapeTree.Append(picture1);

        P14.CreationId creationId1 = new() { Val = (UInt32Value)3972997422U };
        creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        return new AudioShape(this.shapeTree, this.slideObject);
    }

    public IVideoShape AddVideo(int x, int y, Stream stream)
    {
        long xEmu = UnitConverter.HorizontalPixelToEmu(x);
        long yEmu = UnitConverter.VerticalPixelToEmu(y);

        var slideBase = this.slideObject.Match(slide => slide as SlideObject, layout => layout, master => master);
        MediaDataPart mediaDataPart =
            slideBase.PresentationInternal.SDKPresentationInternal.CreateMediaDataPart("video/mp4", ".mp4");

        stream.Position = 0;
        mediaDataPart.FeedData(stream);
        string imgPartRId = $"rId{Guid.NewGuid().ToString().Replace("-", string.Empty).Substring(0, 5)}";
        var slidePart = (SlidePart)slideBase.TypedOpenXmlPart;
        var imagePart = slidePart.AddNewPart<ImagePart>("image/png", imgPartRId);
        var imageStream = Assembly.GetExecutingAssembly().GetStream("video-image.bmp");
        imagePart.FeedData(imageStream);

        var videoRr = slidePart.AddVideoReferenceRelationship(mediaDataPart);
        var mediaRr = slidePart.AddMediaReferenceRelationship(mediaDataPart);

        P.Picture picture1 = new();

        P.NonVisualPictureProperties nonVisualPictureProperties1 = new();

        uint shapeId = (uint)this.CollectionItems.Max(sp => sp.Id) + 1;
        P.NonVisualDrawingProperties nonVisualDrawingProperties2 = new() { Id = shapeId, Name = $"Video{shapeId}" };
        A.HyperlinkOnClick hyperlinkOnClick1 = new A.HyperlinkOnClick() { Id = "", Action = "ppaction://media" };

        A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new();

        A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 =
            new() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

        OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement(
            "<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{2FF36D28-5328-4DA3-BF85-A2B65D7EE127}\" />");

        nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

        nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

        nonVisualDrawingProperties2.Append(hyperlinkOnClick1);
        nonVisualDrawingProperties2.Append(nonVisualDrawingPropertiesExtensionList1);

        P.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new();
        A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

        nonVisualPictureDrawingProperties1.Append(pictureLocks1);

        P.ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new();
        A.VideoFromFile videoFromFile1 = new A.VideoFromFile() { Link = videoRr.Id };

        P.ApplicationNonVisualDrawingPropertiesExtensionList
            applicationNonVisualDrawingPropertiesExtensionList1 = new();

        P.ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 =
            new() { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };

        P14.Media media1 = new P14.Media() { Embed = mediaRr.Id };
        media1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        applicationNonVisualDrawingPropertiesExtension1.Append(media1);

        applicationNonVisualDrawingPropertiesExtensionList1.Append(applicationNonVisualDrawingPropertiesExtension1);

        applicationNonVisualDrawingProperties2.Append(videoFromFile1);
        applicationNonVisualDrawingProperties2.Append(applicationNonVisualDrawingPropertiesExtensionList1);

        nonVisualPictureProperties1.Append(nonVisualDrawingProperties2);
        nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
        nonVisualPictureProperties1.Append(applicationNonVisualDrawingProperties2);

        P.BlipFill blipFill1 = new ();
        A.Blip blip1 = new() { Embed = imgPartRId };

        A.Stretch stretch1 = new ();
        A.FillRectangle fillRectangle1 = new ();

        stretch1.Append(fillRectangle1);

        blipFill1.Append(blip1);
        blipFill1.Append(stretch1);

        P.ShapeProperties shapeProperties1 = new();

        A.Transform2D transform2D1 = new();
        A.Offset offset2 = new() { X = xEmu, Y = yEmu };
        A.Extents extents2 = new() { Cx = 609600L, Cy = 609600L };

        transform2D1.Append(offset2);
        transform2D1.Append(extents2);

        A.PresetGeometry presetGeometry1 = new() { Preset = A.ShapeTypeValues.Rectangle };
        A.AdjustValueList adjustValueList1 = new();

        presetGeometry1.Append(adjustValueList1);

        shapeProperties1.Append(transform2D1);
        shapeProperties1.Append(presetGeometry1);

        picture1.Append(nonVisualPictureProperties1);
        picture1.Append(blipFill1);
        picture1.Append(shapeProperties1);

        this.shapeTree.Append(picture1);

        P14.CreationId creationId1 = new() { Val = (UInt32Value)3972997422U };
        creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        return new VideoShape(this.slideObject, this.shapeTree);
    }

    public IAutoShape AddAutoShape(SCAutoShapeType type, int x, int y, int width, int height)
    {
        var idAndName = this.GenerateIdAndName();

        var adjustValueList = new A.AdjustValueList();
        var presetGeometry = new A.PresetGeometry(adjustValueList) { Preset = A.ShapeTypeValues.Rectangle };
        var shapeProperties = new P.ShapeProperties();
        var xEmu = UnitConverter.HorizontalPixelToEmu(x);
        var yEmu = UnitConverter.VerticalPixelToEmu(y);
        var widthEmu = UnitConverter.HorizontalPixelToEmu(width);
        var heightEmu = UnitConverter.VerticalPixelToEmu(height);
        shapeProperties.AddAXfrm(xEmu, yEmu, widthEmu, heightEmu);
        shapeProperties.Append(presetGeometry);
        
        var newPShape = new P.Shape(
            new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties { Id = (uint)idAndName.Item1, Name = idAndName.Item2 },
            new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
            new ApplicationNonVisualDrawingProperties()),
            shapeProperties,
            new P.TextBody(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph(new A.EndParagraphRunProperties { Language = "en-US" })));

        this.shapeTree.Append(newPShape);
        
        var autoShape = new AutoShape(newPShape, this.slideObject, null);

        return autoShape;
    }

    public ITable AddTable(int xPx, int yPx, int columns, int rows)
    {
        var shapeName = this.GenerateNextTableName();
        var shapeId = this.GenerateNextShapeId();
        var xEmu = UnitConverter.HorizontalPixelToEmu(xPx);
        var yEmu = UnitConverter.VerticalPixelToEmu(yPx);
        var tableHeightEmu = SCConstants.DefaultRowHeightEmu * rows;

        var graphicFrame = new GraphicFrame();
        var nonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties();
        var nonVisualDrawingProperties = new NonVisualDrawingProperties { Id = (uint)shapeId, Name = shapeName };
        var nonVisualGraphicFrameDrawingProperties = new NonVisualGraphicFrameDrawingProperties();
        var applicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();
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
        var tableStyleId = new A.TableStyleId { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" };
        tableProperties.Append(tableStyleId);

        var tableGrid = new A.TableGrid();
        var gridWidthEmu = DefaultTableWidthEmu / columns;
        for (var i = 0; i < columns; i++)
        {
            var gridColumn = new A.GridColumn { Width = gridWidthEmu };
            tableGrid.Append(gridColumn);
        }

        aTable.Append(tableProperties);
        aTable.Append(tableGrid);
        for (var i = 0; i < rows; i++)
        {
            aTable.AddRow(columns);
        }
        
        graphicData.Append(aTable);
        graphic.Append(graphicData);
        graphicFrame.Append(nonVisualGraphicFrameProperties);
        graphicFrame.Append(pTransform);
        graphicFrame.Append(graphic);

        this.shapeTree.Append(graphicFrame);
        var table = new SCTable(graphicFrame, this.slideObject, null);

        return table;
    }

    public void Remove(IShape shape)
    {
        this.CollectionItems.Remove(shape);
        
        var shapeInternal = (Shape)shape;
        shapeInternal.PShapeTreesChild.Remove();
    }

    public T? GetById<T>(int shapeId)
        where T : IShape
    {
           var shape = this.CollectionItems.FirstOrDefault(shape => shape.Id == shapeId);
           return (T?)shape;
    }

    public T? GetByName<T>(string shapeName)
        where T : IShape
    {
        var shape = this.GetByName(shapeName);
  
        return (T?)shape;
    }
    
    public IShape? GetByName(string shapeName)
    {
        return this.CollectionItems.FirstOrDefault(shape => shape.Name == shapeName);
    }

    public Shape? GetReferencedShapeOrNull(P.PlaceholderShape inputPph)
    {
        var phShapes = this.CollectionItems.Where(sp => sp.Placeholder != null).OfType<Shape>();
        var referencedShape = phShapes.FirstOrDefault(IsEqual);

        // https://answers.microsoft.com/en-us/msoffice/forum/all/placeholder-master/0d51dcec-f982-4098-b6b6-94785304607a?page=3
        if (referencedShape == null && inputPph.Index?.Value == 4294967295 && this.slideObject.IsT2)
        {
            var custom = phShapes.Select(sp =>
            {
                var placeholder = (Placeholder?)sp.Placeholder;
                return new
                {
                    shape = sp,
                    index = placeholder?.PPlaceholderShape.Index?.Value
                };
            });

            return custom.FirstOrDefault(x => x.index == 1)?.shape;
        }

        return referencedShape;

        bool IsEqual(Shape collectionShape)
        {
            var placeholder = (Placeholder)collectionShape.Placeholder!;
            var pPh = placeholder.PPlaceholderShape;

            if (inputPph.Index is not null && pPh.Index is not null &&
                inputPph.Index == pPh.Index)
            {
                return true;
            }

            if (inputPph.Type == null || pPh.Type == null)
            {
                return false;
            }

            if (inputPph.Type == P.PlaceholderValues.Body &&
                inputPph.Index is not null && pPh.Index is not null)
            {
                return inputPph.Index == pPh.Index;
            }

            var left = inputPph.Type;
            if (inputPph.Type == PlaceholderValues.CenteredTitle)
            {
                left = PlaceholderValues.Title;
            }

            var right = pPh.Type;
            if (pPh.Type == PlaceholderValues.CenteredTitle)
            {
                right = PlaceholderValues.Title;
            }

            return left.Equals(right);
        }
    }

    internal static ShapeCollection Create(
        OneOf<SlidePart, SlideLayoutPart, SlideMasterPart> oneOfSlidePart,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide)
    {
        var chartGrFrameHandler = new ChartGraphicFrameHandler();
        var tableGrFrameHandler = new TableGraphicFrameHandler();
        var oleGrFrameHandler = new OleGraphicFrameHandler();
        var autoShapeCreator = new AutoShapeCreator();
        var pictureHandler = new PictureHandler();

        autoShapeCreator.Successor = oleGrFrameHandler;
        oleGrFrameHandler.Successor = pictureHandler;
        pictureHandler.Successor = chartGrFrameHandler;
        chartGrFrameHandler.Successor = tableGrFrameHandler;

        var pShapeTree = oneOfSlidePart.Match(
            slidePart => slidePart.Slide.CommonSlideData!.ShapeTree!,
            layoutPart => layoutPart.SlideLayout.CommonSlideData!.ShapeTree!,
            masterPart => masterPart.SlideMaster.CommonSlideData!.ShapeTree!);
        var shapes = new List<IShape>(pShapeTree.Count());
        foreach (var childElementOfShapeTree in pShapeTree.OfType<OpenXmlCompositeElement>())
        {
            IShape? shape;
            if (childElementOfShapeTree is P.GroupShape pGroupShape)
            {
                shape = new SCGroupShape(pGroupShape, oneOfSlide, null);
            }
            else if (childElementOfShapeTree is P.ConnectionShape)
            {
                shape = new SCConnectionShape(childElementOfShapeTree, oneOfSlide);
            }
            else
            {
                shape = autoShapeCreator.Create(childElementOfShapeTree, oneOfSlide, null);
            }

            if (shape != null)
            {
                shapes.Add(shape);
            }
        }

        return new ShapeCollection(shapes, pShapeTree, oneOfSlide);
    }
    
    private int GenerateNextShapeId()
    {
        var maxId = this.CollectionItems.Select(shape => shape.Id).Prepend(0).Max();
        
        return maxId + 1;
    }
    
    
    private string GenerateNextTableName()
    {
        var maxOrder = 0;
        foreach (var shape in this.CollectionItems)
        {
            var matchOrder = Regex.Match(shape.Name, "(?!Table )\\d+");
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
    
    private (int, string) GenerateIdAndName()
    {
        var maxOrder = 0;
        var maxId = 0;
        foreach (var shape in this.CollectionItems)
        {
            if (shape.Id > maxId)
            {
                maxId = shape.Id;
            }

            var matchOrder = Regex.Match(shape.Name, "(?!AutoShape )\\d+");
            if (matchOrder.Success)
            {
                var order = int.Parse(matchOrder.Value);
                if (order > maxOrder)
                {
                    maxOrder = order;
                }
            }
        }

        var shapeId = maxId + 1;
        var shapeName = $"AutoShape {maxOrder + 1}";
        
        return (shapeId, shapeName);
    }
}