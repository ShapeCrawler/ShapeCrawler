using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OneOf;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Constants;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a shape collection.
/// </summary>
public interface IShapeCollection : IReadOnlyList<IShape>
{
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
    ///     Adds a new audio from stream.
    /// </summary>
    /// <param name="xPixel">The X coordinate for the left side of the shape.</param>
    /// <param name="yPixels">The Y coordinate for the left side of the shape.</param>
    /// <param name="mp3Stream">Audio stream data.</param>
    IAudioShape AddAudio(int xPixel, int yPixels, Stream mp3Stream);

    /// <summary>
    ///     Adds a new video from stream.
    /// </summary>
    /// <param name="x">X coordinate in pixels.</param>
    /// <param name="y">Y coordinate in pixels.</param>
    /// <param name="stream">Video stream data.</param>
    IVideoShape AddVideo(int x, int y, Stream stream);

    /// <summary>
    ///     Adds a new Rectangle shape.
    /// </summary>
    IRectangle AddRectangle(int x, int y, int w, int h);

    /// <summary>
    ///     Adds a new Rounded Rectangle shape. 
    /// </summary>
    IRoundedRectangle AddRoundedRectangle(int x, int y, int w, int h);

    /// <summary>
    ///     Adds a line from XML.
    /// </summary>
    /// <param name="xml">Content of p:cxnSp Open XML element.</param>
    ILine AddLine(string xml);

    /// <summary>
    ///     Adds a new line.
    /// </summary>
    ILine AddLine(int startPointX, int startPointY, int endPointX, int endPointY);

    /// <summary>
    ///     Adds a new table.
    /// </summary>
    ITable AddTable(int x, int y, int columnsCount, int rowsCount);

    /// <summary>
    ///     Removes specified shape.
    /// </summary>
    void Remove(IShape shape);
}

internal sealed class ShapeCollection : IShapeCollection
{
    private const long DefaultTableWidthEmu = 8128000L;
    private readonly P.ShapeTree pShapeTree;
    private readonly ResettableLazy<List<IShape>> shapes;

    internal ShapeCollection(
        OneOf<SlidePart, SlideLayoutPart, SlideMasterPart> parentSlidePartOf,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> parentSlideStructureOf)
    {
        this.ParentSlideStructure = parentSlideStructureOf;

        var chartGrFrameHandler = new ChartGraphicFrameHandler();
        var tableGrFrameHandler = new TableGraphicFrameHandler();
        var oleGrFrameHandler = new OleGraphicFrameHandler();
        var autoShapeCreator = new AutoShapeCreator();
        var pictureHandler = new PictureHandler();

        autoShapeCreator.Successor = oleGrFrameHandler;
        oleGrFrameHandler.Successor = pictureHandler;
        pictureHandler.Successor = chartGrFrameHandler;
        chartGrFrameHandler.Successor = tableGrFrameHandler;

        this.pShapeTree = parentSlidePartOf.Match(
            slidePart => slidePart.Slide.CommonSlideData!.ShapeTree!,
            layoutPart => layoutPart.SlideLayout.CommonSlideData!.ShapeTree!,
            masterPart => masterPart.SlideMaster.CommonSlideData!.ShapeTree!);

        this.shapes = new ResettableLazy<List<IShape>>(() => this.GetShapes(autoShapeCreator));
    }

    public int Count => this.shapes.Value.Count;

    internal OneOf<SCSlide, SCSlideLayout, SCSlideMaster> ParentSlideStructure { get; }

    public IShape this[int index] => this.shapes.Value[index];

    public IAudioShape AddAudio(int xPixels, int yPixels, Stream mp3Stream)
    {
        var xEmu = UnitConverter.HorizontalPixelToEmu(xPixels);
        var yEmu = UnitConverter.VerticalPixelToEmu(yPixels);

        var slideBase =
            this.ParentSlideStructure.Match(slide => slide as SlideStructure, layout => layout, master => master);
        var mediaDataPart =
            slideBase.PresentationInternal.SDKPresentationInternal.CreateMediaDataPart("audio/mpeg", ".mp3");

        mp3Stream.Position = 0;
        mediaDataPart.FeedData(mp3Stream);
        var imgPartRId = $"rId{Guid.NewGuid().ToString().Replace("-", string.Empty).Substring(0, 5)}";
        var slidePart = (SlidePart)slideBase.TypedOpenXmlPart;
        var imagePart = slidePart!.AddNewPart<ImagePart>("image/png", imgPartRId);
        var imgStream = Assembly.GetExecutingAssembly().GetStream("audio-image.png");
        imgStream.Position = 0;
        imagePart.FeedData(imgStream);

        var audioRr = slidePart.AddAudioReferenceRelationship(mediaDataPart);
        var mediaRr = slidePart.AddMediaReferenceRelationship(mediaDataPart);

        P.Picture picture1 = new();

        P.NonVisualPictureProperties nonVisualPictureProperties1 = new();

        var shapeId = (uint)this.shapes.Value.Max(sp => sp.Id) + 1;
        P.NonVisualDrawingProperties nonVisualDrawingProperties2 = new() { Id = shapeId, Name = $"Audio{shapeId}" };
        var hyperlinkOnClick1 = new A.HyperlinkOnClick()
            { Id = string.Empty, Action = "ppaction://media" };

        A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new();

        A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 =
            new() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

        nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

        nonVisualDrawingProperties2.Append(hyperlinkOnClick1);
        nonVisualDrawingProperties2.Append(nonVisualDrawingPropertiesExtensionList1);

        P.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new();
        var pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

        nonVisualPictureDrawingProperties1.Append(pictureLocks1);

        P.ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new();
        var audioFromFile1 = new A.AudioFromFile() { Link = audioRr.Id };

        P.ApplicationNonVisualDrawingPropertiesExtensionList
            applicationNonVisualDrawingPropertiesExtensionList1 = new();

        P.ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 =
            new() { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };

        var media1 = new P14.Media() { Embed = mediaRr.Id };
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

        this.pShapeTree.Append(picture1);

        P14.CreationId creationId1 = new() { Val = (UInt32Value)3972997422U };
        creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        return new SCAudioShape(this.pShapeTree, this.ParentSlideStructure, this);
    }

    public IVideoShape AddVideo(int x, int y, Stream stream)
    {
        var xEmu = UnitConverter.HorizontalPixelToEmu(x);
        var yEmu = UnitConverter.VerticalPixelToEmu(y);

        var slideBase =
            this.ParentSlideStructure.Match(slide => slide as SlideStructure, layout => layout, master => master);
        var mediaDataPart =
            slideBase.PresentationInternal.SDKPresentationInternal.CreateMediaDataPart("video/mp4", ".mp4");

        stream.Position = 0;
        mediaDataPart.FeedData(stream);
        var imgPartRId = $"rId{Guid.NewGuid().ToString().Replace("-", string.Empty).Substring(0, 5)}";
        var slidePart = (SlidePart)slideBase.TypedOpenXmlPart;
        var imagePart = slidePart.AddNewPart<ImagePart>("image/png", imgPartRId);
        var imageStream = Assembly.GetExecutingAssembly().GetStream("video-image.bmp");
        imagePart.FeedData(imageStream);

        var videoRr = slidePart.AddVideoReferenceRelationship(mediaDataPart);
        var mediaRr = slidePart.AddMediaReferenceRelationship(mediaDataPart);

        P.Picture picture1 = new();

        P.NonVisualPictureProperties nonVisualPictureProperties1 = new();

        var shapeId = (uint)this.shapes.Value.Max(sp => sp.Id) + 1;
        P.NonVisualDrawingProperties nonVisualDrawingProperties2 = new() { Id = shapeId, Name = $"Video{shapeId}" };
        var hyperlinkOnClick1 = new A.HyperlinkOnClick()
            { Id = string.Empty, Action = "ppaction://media" };

        A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new();

        A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 =
            new() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

        nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

        nonVisualDrawingProperties2.Append(hyperlinkOnClick1);
        nonVisualDrawingProperties2.Append(nonVisualDrawingPropertiesExtensionList1);

        P.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new();
        var pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

        nonVisualPictureDrawingProperties1.Append(pictureLocks1);

        P.ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new();
        var videoFromFile1 = new A.VideoFromFile() { Link = videoRr.Id };

        P.ApplicationNonVisualDrawingPropertiesExtensionList
            applicationNonVisualDrawingPropertiesExtensionList1 = new();

        P.ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 =
            new() { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };

        var media1 = new P14.Media() { Embed = mediaRr.Id };
        media1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        applicationNonVisualDrawingPropertiesExtension1.Append(media1);

        applicationNonVisualDrawingPropertiesExtensionList1.Append(applicationNonVisualDrawingPropertiesExtension1);

        applicationNonVisualDrawingProperties2.Append(videoFromFile1);
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

        this.pShapeTree.Append(picture1);

        P14.CreationId creationId1 = new() { Val = (UInt32Value)3972997422U };
        creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        return new SCVideoShape(this.pShapeTree, this.ParentSlideStructure, this);
    }

    public IRectangle AddRectangle(int x, int y, int width, int height)
    {
        var newPShape = this.CreatePShape(x, y, width, height, A.ShapeTypeValues.Rectangle);

        var newShape = new SCRectangle(newPShape, this.ParentSlideStructure, this);
        newShape.Outline.Color = "000000";

        newShape.Duplicated += this.OnAutoShapeAdded;
        this.shapes.Value.Add(newShape);
        this.pShapeTree.Append(newPShape);

        return newShape;
    }

    public IRoundedRectangle AddRoundedRectangle(int x, int y, int w, int h)
    {
        var newPShape = this.CreatePShape(x, y, w, h, A.ShapeTypeValues.RoundRectangle);

        var newShape = new SCRoundedRectangle(newPShape, this.ParentSlideStructure, this);
        newShape.Outline.Color = "000000";

        newShape.Duplicated += this.OnAutoShapeAdded;
        this.shapes.Value.Add(newShape);
        this.pShapeTree.Append(newPShape);

        return newShape;
    }

    public ILine AddLine(string xml)
    {
        var newPConnectionShape = new ConnectionShape(xml);

        var newShape = new SCLine(newPConnectionShape, this.ParentSlideStructure, this);

        newShape.Duplicated += this.OnAutoShapeAdded;
        this.shapes.Value.Add(newShape);
        this.pShapeTree.Append(newPConnectionShape);

        return newShape;
    }

    public ILine AddLine(int startPointX, int startPointY, int endPointX, int endPointY)
    {
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

        if(startPointX > endPointX)
        {
            flipH = true;
        }
        
        var newPConnectionShape = this.CreatePConnectionShape(x, y, (int)cx, cy, flipH, flipV);

        var newShape = new SCLine(newPConnectionShape, this.ParentSlideStructure, this);
        newShape.Outline.Color = "000000";

        newShape.Duplicated += this.OnAutoShapeAdded;
        this.shapes.Value.Add(newShape);
        this.pShapeTree.Append(newPConnectionShape);

        return newShape;
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

        this.pShapeTree.Append(graphicFrame);
        var table = new SCTable(graphicFrame, this.ParentSlideStructure, this);

        return table;
    }

    public void Remove(IShape shape)
    {
        this.shapes.Value.Remove(shape);

        var shapeInternal = (SCShape)shape;
        shapeInternal.PShapeTreeChild.Remove();
    }

    public T? GetById<T>(int shapeId)
        where T : IShape
    {
        var shape = this.shapes.Value.FirstOrDefault(shape => shape.Id == shapeId);
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
        return this.shapes.Value.FirstOrDefault(shape => shape.Name == shapeName);
    }

    public SCShape? GetReferencedShapeOrNull(P.PlaceholderShape inputPph)
    {
        var phShapes = this.shapes.Value.Where(sp => sp.Placeholder != null).OfType<SCShape>();
        var referencedShape = phShapes.FirstOrDefault(IsEqual);

        // https://answers.microsoft.com/en-us/msoffice/forum/all/placeholder-master/0d51dcec-f982-4098-b6b6-94785304607a?page=3
        if (referencedShape == null && inputPph.Index?.Value == 4294967295 && this.ParentSlideStructure.IsT2)
        {
            var custom = phShapes.Select(sp =>
            {
                var placeholder = (SCPlaceholder?)sp.Placeholder;
                return new
                {
                    shape = sp,
                    index = placeholder?.PPlaceholderShape.Index?.Value
                };
            });

            return custom.FirstOrDefault(x => x.index == 1)?.shape;
        }

        return referencedShape;

        bool IsEqual(SCShape collectionShape)
        {
            var placeholder = (SCPlaceholder)collectionShape.Placeholder!;
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

    public IEnumerator<IShape> GetEnumerator()
    {
        return this.shapes.Value.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    private P.Shape CreatePShape(int x, int y, int width, int height, A.ShapeTypeValues form)
    {
        var idAndName = this.GenerateIdAndName();
        var adjustValueList = new A.AdjustValueList();
        var presetGeometry = new A.PresetGeometry(adjustValueList) { Preset = form };
        var shapeProperties = new P.ShapeProperties();
        var xEmu = UnitConverter.HorizontalPixelToEmu(x);
        var yEmu = UnitConverter.VerticalPixelToEmu(y);
        var widthEmu = UnitConverter.HorizontalPixelToEmu(width);
        var heightEmu = UnitConverter.VerticalPixelToEmu(height);
        shapeProperties.AddAXfrm(xEmu, yEmu, widthEmu, heightEmu);
        shapeProperties.Append(presetGeometry);

        var aRunProperties = new A.RunProperties { Language = "en-US" };
        var aText = new A.Text(string.Empty);
        var aRun = new A.Run(aRunProperties, aText);
        var aEndParaRPr = new A.EndParagraphRunProperties { Language = "en-US" };
        var aParagraph = new A.Paragraph(aRun, aEndParaRPr);

        var pShape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = (uint)idAndName.Item1, Name = idAndName.Item2 },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            shapeProperties,
            new P.TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                aParagraph));

        return pShape;
    }

    private P.ConnectionShape CreatePConnectionShape(int xPx, int yPx, int cxPx, int cyPx, bool flipH, bool flipV)
    {
        var idAndName = this.GenerateIdAndName();
        var adjustValueList = new A.AdjustValueList();
        var presetGeometry = new A.PresetGeometry(adjustValueList) { Preset = A.ShapeTypeValues.Line };
        var shapeProperties = new P.ShapeProperties();
        var xEmu = UnitConverter.HorizontalPixelToEmu(xPx);
        var yEmu = UnitConverter.VerticalPixelToEmu(yPx);
        var cxEmu = UnitConverter.HorizontalPixelToEmu(cxPx);
        var cyEmu = UnitConverter.VerticalPixelToEmu(cyPx);
        var aXfrm = shapeProperties.AddAXfrm(xEmu, yEmu, cxEmu, cyEmu);
        aXfrm.HorizontalFlip = new BooleanValue(flipH);
        aXfrm.VerticalFlip = new BooleanValue(flipV);
        shapeProperties.Append(presetGeometry);

        var pConnectionShape = new P.ConnectionShape(
            new P.NonVisualConnectionShapeProperties(
                new P.NonVisualDrawingProperties { Id = (uint)idAndName.Item1, Name = idAndName.Item2 },
                new P.NonVisualConnectorShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            shapeProperties);

        return pConnectionShape;
    }

    private (int, string) GenerateIdAndName()
    {
        var maxId = 0;
        if (this.shapes.Value.Any())
        {
            maxId = this.shapes.Value.Max(s => s.Id);
        }

        var maxOrder = Regex.Matches(string.Join(string.Empty, this.shapes.Value.Select(s => s.Name)), "\\d+")
#if NETSTANDARD2_0
            .Cast<Match>()
#endif
            .Select(m => int.Parse(m.Value))
            .DefaultIfEmpty(0)
            .Max();

        return (maxId + 1, $"AutoShape {maxOrder + 1}");
    }

    private int GenerateNextShapeId()
    {
        return this.shapes.Value.Select(shape => shape.Id).Prepend(0).Max() + 1;
    }

    private void OnAutoShapeAdded(object? sender, NewAutoShape newAutoShape)
    {
        this.pShapeTree.Append(newAutoShape.PShapeTreeChild);
        newAutoShape.AutoShape.Duplicated += this.OnAutoShapeAdded;

        this.shapes.Reset();
    }

    private string GenerateNextTableName()
    {
        var maxOrder = 0;
        foreach (var shape in this.shapes.Value)
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

    private List<IShape> GetShapes(AutoShapeCreator autoShapeCreator)
    {
        var shapesValue = new List<IShape>(this.pShapeTree.Count());
        foreach (var pShapeTreeChild in this.pShapeTree.OfType<TypedOpenXmlCompositeElement>())
        {
            IShape? shape;
            if (pShapeTreeChild is P.GroupShape pGroupShape)
            {
                shape = new SCGroupShape(pGroupShape, this.ParentSlideStructure, this);
                shapesValue.Add(shape);
            }
            else if (pShapeTreeChild is P.ConnectionShape)
            {
                shape = new SCLine(pShapeTreeChild, this.ParentSlideStructure, this);
                shapesValue.Add(shape);
            }
            else
            {
                shape = autoShapeCreator.Create(pShapeTreeChild, this.ParentSlideStructure, this);
                if (shape != null)
                {
                    shapesValue.Add(shape);
                }

                if (shape is SCAutoShape autoShape)
                {
                    autoShape.Duplicated += this.OnAutoShapeAdded;
                }
            }
        }

        return shapesValue;
    }
}