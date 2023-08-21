using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Charts;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Services;
using ShapeCrawler.Shared;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Shapes;

internal sealed class SlideShapes : ISlideShapeCollection
{
    private const long DefaultTableWidthEmu = 8128000L;
    private readonly P.ShapeTree pShapeTree;
    private readonly ResetableLazy<List<IShape>> shapes;
    private readonly SCSlide parentSlide;

    internal SlideShapes(SCSlide parentSlide, P.ShapeTree pShapeTree)
    {
        this.parentSlide = parentSlide;
        this.pShapeTree = pShapeTree;
        this.shapes = new ResetableLazy<List<IShape>>(this.ParseShapes);
    }
    
    internal void Add(OpenXmlElement pShapeTreeChild)
    {
        var nextId = 1;
        if (this.shapes.Value.Any())
        {
            nextId = this.shapes.Value.Select(shape => shape.Id).Prepend(0).Max() + 1;
        }
        
        pShapeTreeChild.GetNonVisualDrawingProperties().Id = new UInt32Value((uint)nextId);
        this.pShapeTree.Append(pShapeTreeChild);
        
        this.shapes.Reset();
    }

    public int Count => this.shapes.Value.Count;

    public IShape this[int index] => this.shapes.Value[index];

    internal int CalculateNextShapeId()
    {
        if (this.shapes.Value.Any())
        {
            return this.shapes.Value.Select(shape => shape.Id).Prepend(0).Max() + 1;
        }

        return 1;
    }

    public void Add(IShape addingShape)
    {
        var id = this.CalculateNextShapeId();
        var allShapeNames = this.Select(shape => shape.Name);
        
        if (addingShape is SlideAutoShape slideAutoShape)
        {
            slideAutoShape.CopyTo(id, this.pShapeTree, allShapeNames, this.parentSlide.SDKSlidePart());    
        }
        else if (addingShape is SCSlidePicture slidePicture)
        {
            slidePicture.CopyTo(id, this.pShapeTree, allShapeNames, this.parentSlide.SDKSlidePart());    
        }
        else
        {
            throw new SCException($"Adding {addingShape.GetType().Name} is not supported.");
        }

        this.shapes.Reset();
    }

    public IMediaShape AddAudio(int xPixels, int yPixels, Stream mp3Stream)
    {
        var xEmu = UnitConverter.HorizontalPixelToEmu(xPixels);
        var yEmu = UnitConverter.VerticalPixelToEmu(yPixels);
        var mediaDataPart = this.sdkPresentationDocument.CreateMediaDataPart("audio/mpeg", ".mp3");
        mp3Stream.Position = 0;
        mediaDataPart.FeedData(mp3Stream);
        var imageStream = Assembly.GetExecutingAssembly().GetStream("audio-image.png");

        var audioRef = this.sdkSlidePart.AddAudioReferenceRelationship(mediaDataPart);
        var mediaRef = this.sdkSlidePart.AddMediaReferenceRelationship(mediaDataPart);

        var audioFromFile = new DocumentFormat.OpenXml.Drawing.AudioFromFile() { Link = audioRef.Id };

        var appNonVisualDrawingPropsExtensionList = new P.ApplicationNonVisualDrawingPropertiesExtensionList();

        var appNonVisualDrawingPropsExtension = new P.ApplicationNonVisualDrawingPropertiesExtension
            { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };

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
        var nonVisualDrawingProps = pPicture.GetNonVisualDrawingProperties();
        var hyperlinkOnClick = new DocumentFormat.OpenXml.Drawing.HyperlinkOnClick
            { Id = string.Empty, Action = "ppaction://media" };
        nonVisualDrawingProps.Append(hyperlinkOnClick);
        nonVisualPictureProps.Append(new P.NonVisualPictureDrawingProperties());

        var applicationNonVisualDrawingProps = nonVisualPictureProps.ApplicationNonVisualDrawingProperties!;
        applicationNonVisualDrawingProps.Append(audioFromFile);
        applicationNonVisualDrawingProps.Append(appNonVisualDrawingPropsExtensionList);

        this.shapes.Reset();

        return new SlideMediaShape(this.pShapeTree, this.parentSlide, this, pPicture);
    }

    public IPicture AddPicture(Stream imageStream)
    {
        imageStream.Position = 0;
        var imageCopy = new MemoryStream();
        imageStream.CopyTo(imageCopy);
        imageCopy.Position = 0;
        imageStream.Position = 0;
        using var skBitmap = SKBitmap.Decode(imageCopy);
        var xEmu = UnitConverter.HorizontalPixelToEmu(100);
        var yEmu = UnitConverter.VerticalPixelToEmu(100);
        var cxEmu = UnitConverter.HorizontalPixelToEmu(skBitmap.Width);
        var cyEmu = UnitConverter.VerticalEmuToPixel(skBitmap.Height);

        var pPicture = this.CreatePPicture(imageStream, "Picture");

        var transform2D = pPicture.ShapeProperties!.Transform2D!;
        transform2D.Offset!.X = xEmu;
        transform2D.Offset!.Y = yEmu;
        transform2D.Extents!.Cx = cxEmu;
        transform2D.Extents!.Cy = cyEmu;

        var pictureHandler = new PictureHandler();
        var shape = pictureHandler.FromTreeChild(pPicture, this.slideOf, this) !;

        this.shapes.Reset();

        return (SCSlidePicture)shape;
    }

    public IChart AddBarChart(BarChartType barChartType)
    {
        var chartFactory = new ChartGraphicFrameHandler();
        var newPGraphicFrame = chartFactory.Create(this.slideTypedOpenXmlPart);

        this.pShapeTree.Append(newPGraphicFrame);

        return null!;
    }

    public IMediaShape AddVideo(int x, int y, Stream stream)
    {
        var xEmu = UnitConverter.HorizontalPixelToEmu(x);
        var yEmu = UnitConverter.VerticalPixelToEmu(y);

        var mediaDataPart =
            this.parentSlide.Presentation.SDKPresentationDocument.CreateMediaDataPart("video/mp4", ".mp4");

        stream.Position = 0;
        mediaDataPart.FeedData(stream);
        var imgPartRId = $"rId{Guid.NewGuid().ToString().Replace("-", string.Empty).Substring(0, 5)}";
        var imagePart = this.slideTypedOpenXmlPart.AddNewPart<ImagePart>("image/png", imgPartRId);
        var imageStream = Assembly.GetExecutingAssembly().GetStream("video-image.bmp");
        imagePart.FeedData(imageStream);

        var slidePart = (SlidePart)this.slideTypedOpenXmlPart;
        var videoRr = slidePart.AddVideoReferenceRelationship(mediaDataPart);
        var mediaRr = slidePart.AddMediaReferenceRelationship(mediaDataPart);

        P.Picture picture1 = new();

        P.NonVisualPictureProperties nonVisualPictureProperties1 = new();

        var shapeId = (uint)this.shapes.Value.Max(sp => sp.Id) + 1;
        P.NonVisualDrawingProperties nonVisualDrawingProperties2 = new() { Id = shapeId, Name = $"Video{shapeId}" };
        var hyperlinkOnClick1 = new DocumentFormat.OpenXml.Drawing.HyperlinkOnClick()
            { Id = string.Empty, Action = "ppaction://media" };

        DocumentFormat.OpenXml.Drawing.NonVisualDrawingPropertiesExtensionList
            nonVisualDrawingPropertiesExtensionList1 = new();

        DocumentFormat.OpenXml.Drawing.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 =
            new() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

        nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

        nonVisualDrawingProperties2.Append(hyperlinkOnClick1);
        nonVisualDrawingProperties2.Append(nonVisualDrawingPropertiesExtensionList1);

        P.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new();
        var pictureLocks1 = new DocumentFormat.OpenXml.Drawing.PictureLocks() { NoChangeAspect = true };

        nonVisualPictureDrawingProperties1.Append(pictureLocks1);

        P.ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new();
        var videoFromFile1 = new DocumentFormat.OpenXml.Drawing.VideoFromFile() { Link = videoRr.Id };

        P.ApplicationNonVisualDrawingPropertiesExtensionList
            applicationNonVisualDrawingPropertiesExtensionList1 = new();

        P.ApplicationNonVisualDrawingPropertiesExtension applicationNonVisualDrawingPropertiesExtension1 =
            new() { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };

        var media1 = new DocumentFormat.OpenXml.Office2010.PowerPoint.Media() { Embed = mediaRr.Id };
        media1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        applicationNonVisualDrawingPropertiesExtension1.Append(media1);

        applicationNonVisualDrawingPropertiesExtensionList1.Append(applicationNonVisualDrawingPropertiesExtension1);

        applicationNonVisualDrawingProperties2.Append(videoFromFile1);
        applicationNonVisualDrawingProperties2.Append(applicationNonVisualDrawingPropertiesExtensionList1);

        nonVisualPictureProperties1.Append(nonVisualDrawingProperties2);
        nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
        nonVisualPictureProperties1.Append(applicationNonVisualDrawingProperties2);

        P.BlipFill blipFill1 = new();
        DocumentFormat.OpenXml.Drawing.Blip blip1 = new() { Embed = imgPartRId };

        DocumentFormat.OpenXml.Drawing.Stretch stretch1 = new();
        DocumentFormat.OpenXml.Drawing.FillRectangle fillRectangle1 = new();

        stretch1.Append(fillRectangle1);

        blipFill1.Append(blip1);
        blipFill1.Append(stretch1);

        P.ShapeProperties shapeProperties1 = new();

        DocumentFormat.OpenXml.Drawing.Transform2D transform2D1 = new();
        DocumentFormat.OpenXml.Drawing.Offset offset2 = new() { X = xEmu, Y = yEmu };
        DocumentFormat.OpenXml.Drawing.Extents extents2 = new() { Cx = 609600L, Cy = 609600L };

        transform2D1.Append(offset2);
        transform2D1.Append(extents2);

        DocumentFormat.OpenXml.Drawing.PresetGeometry presetGeometry1 = new()
            { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle };
        DocumentFormat.OpenXml.Drawing.AdjustValueList adjustValueList1 = new();

        presetGeometry1.Append(adjustValueList1);

        shapeProperties1.Append(transform2D1);
        shapeProperties1.Append(presetGeometry1);

        picture1.Append(nonVisualPictureProperties1);
        picture1.Append(blipFill1);
        picture1.Append(shapeProperties1);

        this.pShapeTree.Append(picture1);

        DocumentFormat.OpenXml.Office2010.PowerPoint.CreationId creationId1 = new() { Val = (UInt32Value)3972997422U };
        creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");

        this.shapes.Reset();

        return new SlideMediaShape(this.pShapeTree, this.slideOf, this, this.slideTypedOpenXmlPart);
    }

    public void AddRectangle(int x, int y, int width, int height)
    {
        var newPShape =
            this.CreatePShape(x, y, width, height, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle);

        var newShape = new SCSlideRectangle(newPShape, this.slideOf, this);
        newShape.Outline.Color = "000000";

        newShape.Duplicated += this.OnAutoShapeAdded;
        this.shapes.Value.Add(newShape);
        this.pShapeTree.Append(newPShape);

        this.shapes.Reset();

        return newShape;
    }

    public void AddRoundedRectangle(int x, int y, int w, int h)
    {
        var newPShape = this.CreatePShape(x, y, w, h, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.RoundRectangle);

        var newShape = new SCRoundedRectangle(newPShape, this.slideOf, this);
        newShape.Outline.Color = "000000";

        newShape.Duplicated += this.OnAutoShapeAdded;
        this.shapes.Value.Add(newShape);
        this.pShapeTree.Append(newPShape);

        this.shapes.Reset();

        return newShape;
    }

    public ILine AddLine(string xml)
    {
        var newPConnectionShape = new ConnectionShape(xml);

        var newShape = new SCLine(newPConnectionShape, this.slideOf, this);

        newShape.Duplicated += this.OnAutoShapeAdded;
        this.shapes.Value.Add(newShape);
        this.pShapeTree.Append(newPConnectionShape);

        this.shapes.Reset();

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

        if (startPointX > endPointX)
        {
            flipH = true;
        }

        var newPConnectionShape = this.CreatePConnectionShape(x, y, (int)cx, cy, flipH, flipV);

        var newShape = new SCLine(newPConnectionShape, this.slideOf, this);
        newShape.Outline.HexColor = "000000";

        newShape.Duplicated += this.OnAutoShapeAdded;
        this.shapes.Value.Add(newShape);
        this.pShapeTree.Append(newPConnectionShape);

        this.shapes.Reset();

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

        var offset = new DocumentFormat.OpenXml.Drawing.Offset { X = xEmu, Y = yEmu };
        var extents = new DocumentFormat.OpenXml.Drawing.Extents { Cx = DefaultTableWidthEmu, Cy = tableHeightEmu };
        var pTransform = new P.Transform(offset, extents);

        var graphic = new DocumentFormat.OpenXml.Drawing.Graphic();
        var graphicData = new DocumentFormat.OpenXml.Drawing.GraphicData
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" };
        var aTable = new DocumentFormat.OpenXml.Drawing.Table();

        var tableProperties = new DocumentFormat.OpenXml.Drawing.TableProperties { FirstRow = true, BandRow = true };
        var tableStyleId = new DocumentFormat.OpenXml.Drawing.TableStyleId
            { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" };
        tableProperties.Append(tableStyleId);

        var tableGrid = new DocumentFormat.OpenXml.Drawing.TableGrid();
        var gridWidthEmu = DefaultTableWidthEmu / columns;
        for (var i = 0; i < columns; i++)
        {
            var gridColumn = new DocumentFormat.OpenXml.Drawing.GridColumn { Width = gridWidthEmu };
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
        var table = new SlideTable(graphicFrame, this.slideOf, this);

        this.shapes.Reset();

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

    public IEnumerator<IShape> GetEnumerator()
    {
        return this.shapes.Value.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    private P.Shape CreatePShape(int x, int y, int width, int height,
        DocumentFormat.OpenXml.Drawing.ShapeTypeValues form)
    {
        var idAndName = this.GenerateIdAndName();
        var adjustValueList = new DocumentFormat.OpenXml.Drawing.AdjustValueList();
        var presetGeometry = new DocumentFormat.OpenXml.Drawing.PresetGeometry(adjustValueList) { Preset = form };
        var shapeProperties = new P.ShapeProperties();
        var xEmu = UnitConverter.HorizontalPixelToEmu(x);
        var yEmu = UnitConverter.VerticalPixelToEmu(y);
        var widthEmu = UnitConverter.HorizontalPixelToEmu(width);
        var heightEmu = UnitConverter.VerticalPixelToEmu(height);
        shapeProperties.AddAXfrm(xEmu, yEmu, widthEmu, heightEmu);
        shapeProperties.Append(presetGeometry);

        var aRunProperties = new DocumentFormat.OpenXml.Drawing.RunProperties { Language = "en-US" };
        var aText = new DocumentFormat.OpenXml.Drawing.Text(string.Empty);
        var aRun = new DocumentFormat.OpenXml.Drawing.Run(aRunProperties, aText);
        var aEndParaRPr = new DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties { Language = "en-US" };
        var aParagraph = new DocumentFormat.OpenXml.Drawing.Paragraph(aRun, aEndParaRPr);

        var pShape = new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties
                {
                    Id = (uint)idAndName.Item1, Name = idAndName.Item2
                },
                new P.NonVisualShapeDrawingProperties(new DocumentFormat.OpenXml.Drawing.ShapeLocks
                    { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            shapeProperties,
            new P.TextBody(
                new DocumentFormat.OpenXml.Drawing.BodyProperties(),
                new DocumentFormat.OpenXml.Drawing.ListStyle(),
                aParagraph));

        return pShape;
    }

    private P.ConnectionShape CreatePConnectionShape(int xPx, int yPx, int cxPx, int cyPx, bool flipH, bool flipV)
    {
        var idAndName = this.GenerateIdAndName();
        var adjustValueList = new DocumentFormat.OpenXml.Drawing.AdjustValueList();
        var presetGeometry = new DocumentFormat.OpenXml.Drawing.PresetGeometry(adjustValueList)
            { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Line };
        var shapeProperties = new P.ShapeProperties();
        var xEmu = UnitConverter.HorizontalPixelToEmu(xPx);
        var yEmu = UnitConverter.VerticalPixelToEmu(yPx);
        var cxEmu = UnitConverter.HorizontalPixelToEmu(cxPx);
        var cyEmu = UnitConverter.VerticalPixelToEmu(cyPx);
        var aXfrm = shapeProperties.AddAXfrm(xEmu, yEmu, cxEmu, cyEmu);
        aXfrm.HorizontalFlip = new BooleanValue(flipH);
        aXfrm.VerticalFlip = new BooleanValue(flipV);
        shapeProperties.Append(presetGeometry);

        var pConnectionShape = new ConnectionShape(
            new NonVisualConnectionShapeProperties(
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

        var maxOrder = Regex.Matches(string.Join(string.Empty, this.shapes.Value.Select(s => s.Name)), "\\d+",
                RegexOptions.None, TimeSpan.FromSeconds(100))

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

    private List<IShape> ParseShapes()
    {
        var shapesValue = new List<IShape>(this.pShapeTree.Count());
        foreach (var pShapeTreeChild in this.pShapeTree.OfType<TypedOpenXmlCompositeElement>())
        {
            if (pShapeTreeChild is P.GroupShape pGroupShape)
            {
                shapesValue.Add(new SCSlideGroupShape(pGroupShape, this));
            }
            else if (pShapeTreeChild is P.ConnectionShape pConnectionShape)
            {
                shapesValue.Add(new SCLine(pConnectionShape, this, new Shape(pShapeTreeChild)));
            }
            else if (pShapeTreeChild is P.Shape pShape)
            {
                shapesValue.Add(new SlideAutoShape(pShape, this, new Shape(pShapeTreeChild)));
            }
            else if (pShapeTreeChild is P.GraphicFrame pGraphicFrame)
            {
                var aGraphicData = pShapeTreeChild.GetFirstChild<A.Graphic>()!.GetFirstChild<A.GraphicData>();
                if (aGraphicData!.Uri!.Value!.Equals("http://schemas.openxmlformats.org/presentationml/2006/ole",
                        StringComparison.Ordinal))
                {
                    shapesValue.Add(new SCSlideOLEObject(pGraphicFrame, this, new Shape(pShapeTreeChild)));
                    continue;
                }

                var pPicture = pShapeTreeChild.Descendants<P.Picture>().FirstOrDefault();
                if (pPicture != null)
                {
                    var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
                    var blipEmbed = aBlip?.Embed;
                    if (blipEmbed is null)
                    {
                        continue;
                    }

                    shapesValue.Add(new SCSlidePicture(pPicture, this, aBlip!, new Shape(pShapeTreeChild)));
                    continue;
                }
            }
            else if (pShapeTreeChild is P.Picture pPicture)
            {
                var element = pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!.ChildElements
                    .FirstOrDefault();

                switch (element)
                {
                    case A.AudioFromFile:
                    {
                        var aAudioFile = pPicture.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties
                            .GetFirstChild<A.AudioFromFile>();
                        if (aAudioFile is not null)
                        {
                            shapesValue.Add(new SlideMediaShape(pPicture, this));
                        }

                        continue;
                    }
                    case A.VideoFromFile:
                    {
                        shapesValue.Add(new SlideMediaShape(pPicture, this));
                        continue;
                    }
                }

                var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
                var blipEmbed = aBlip?.Embed;
                if (blipEmbed is null)
                {
                    continue;
                }

                shapesValue.Add(new SCSlidePicture(pPicture, this, aBlip!, new Shape(pShapeTreeChild)));
                continue;
            }
            else if (this.IsChartPGraphicFrame(pShapeTreeChild))
            {
                var aGraphicData = pShapeTreeChild.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>() !;
                var cChartRef = aGraphicData.GetFirstChild<C.ChartReference>() !;
                var sdkSlidePart = this.parentSlide.SDKSlidePart();
                var chartPart = (ChartPart)sdkSlidePart.GetPartById(cChartRef.Id!);
                var cPlotArea = chartPart!.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea;
                var cCharts = cPlotArea!.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));

                if (cCharts.Count() > 1)
                {
                    shapesValue.Add(new SCSlideChart((P.GraphicFrame)pShapeTreeChild, this));
                    continue;
                }

                var chartTypeName = cCharts.Single().LocalName;

                if (chartTypeName == "lineChart")
                {
                    shapesValue.Add(new SCSlideChart((P.GraphicFrame)pShapeTreeChild, this));
                }

                if (chartTypeName == "barChart")
                {
                    shapesValue.Add(new SCSlideChart((P.GraphicFrame)pShapeTreeChild, this));
                }

                if (chartTypeName == "pieChart")
                {
                    shapesValue.Add(new SCSlideChart((P.GraphicFrame)pShapeTreeChild, this));
                }

                if (chartTypeName == "scatterChart")
                {
                    shapesValue.Add(new SCSlideChart((P.GraphicFrame)pShapeTreeChild, this));
                }

                shapesValue.Add(new SCSlideChart((P.GraphicFrame)pShapeTreeChild, this));
            }
            else if (this.IsTablePGraphicFrame(pShapeTreeChild))
            {
                shapesValue.Add(new SlideTable(pShapeTreeChild, this));
            }
        }

        return shapesValue;
    }

    private bool IsTablePGraphicFrame(TypedOpenXmlCompositeElement pShapeTreeChild)
    {
        if (pShapeTreeChild is P.GraphicFrame pGraphicFrame)
        {
            var graphicData = pGraphicFrame.Graphic!.GraphicData!;
            if (graphicData.Uri!.Value!.Equals("http://schemas.openxmlformats.org/drawingml/2006/table",
                    StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private bool IsChartPGraphicFrame(TypedOpenXmlCompositeElement pShapeTreeChild)
    {
        if (pShapeTreeChild is P.GraphicFrame)
        {
            var aGraphicData = pShapeTreeChild.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>() !;
            if (aGraphicData.Uri!.Value!.Equals("http://schemas.openxmlformats.org/drawingml/2006/chart",
                    StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    private IShape? CreateShape(TypedOpenXmlCompositeElement pShapeTreeChild)
    {
        IShape? shape;

        if (pShapeTreeChild is P.GroupShape pGroupShape)
        {
            return new SCSlideGroupShape(pGroupShape, this.slideOf, this);
        }

        if (pShapeTreeChild is P.ConnectionShape)
        {
            return new SCLine(pShapeTreeChild, this.slideOf, this);
        }

        shape = autoShapeCreator.FromTreeChild(pShapeTreeChild, this.slideOf, this);

        if (shape is SlideAutoShape autoShape)
        {
            autoShape.Duplicated += this.OnAutoShapeAdded;
        }

        return shape;
    }

    private P.Picture CreatePPicture(Stream imageStream, string shapeName)
    {
        var slidePart = this.parentSlide.TypedOpenXmlPart;
        var imgPartRId = slidePart.GetNextRelationshipId();
        var imagePart = slidePart.AddNewPart<ImagePart>("image/png", imgPartRId);
        imageStream.Position = 0;
        imagePart.FeedData(imageStream);

        var nonVisualPictureProperties = new P.NonVisualPictureProperties();
        var shapeId = (uint)this.GenerateNextShapeId();
        var nonVisualDrawingProperties = new P.NonVisualDrawingProperties
            {
                Id = shapeId, Name = $"{shapeName} {shapeId}"
            }
            ;
        var nonVisualPictureDrawingProperties = new P.NonVisualPictureDrawingProperties();
        var appNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties();

        nonVisualPictureProperties.Append(nonVisualDrawingProperties);
        nonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
        nonVisualPictureProperties.Append(appNonVisualDrawingProperties);

        var blipFill = new P.BlipFill();
        var blip = new DocumentFormat.OpenXml.Drawing.Blip { Embed = imgPartRId };
        var stretch = new DocumentFormat.OpenXml.Drawing.Stretch();
        blipFill.Append(blip);
        blipFill.Append(stretch);

        var transform2D = new DocumentFormat.OpenXml.Drawing.Transform2D(
            new DocumentFormat.OpenXml.Drawing.Offset { X = 0, Y = 0 },
            new DocumentFormat.OpenXml.Drawing.Extents { Cx = 0, Cy = 0 });

        var presetGeometry = new DocumentFormat.OpenXml.Drawing.PresetGeometry
            { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle };
        var shapeProperties = new P.ShapeProperties();
        shapeProperties.Append(transform2D);
        shapeProperties.Append(presetGeometry);

        var pPicture = new P.Picture();
        pPicture.Append(nonVisualPictureProperties);
        pPicture.Append(blipFill);
        pPicture.Append(shapeProperties);

        this.pShapeTree.Append(pPicture);

        return pPicture;
    }

    internal SlidePart SDKSlidePart()
    {
        return this.parentSlide.SDKSlidePart();
    }

    public List<ImagePart> SDKImageParts()
    {
        return this.parentSlide.SDKImageParts();
    }

    internal void Reset()
    {
        this.shapes.Reset();
    }

    internal PresentationCore Presentation()
    {
        return this.parentSlide.Presentation();
    }

    internal SlideMaster SlideMaster()
    {
        return this.parentSlide.SlideMaster();
    }
}