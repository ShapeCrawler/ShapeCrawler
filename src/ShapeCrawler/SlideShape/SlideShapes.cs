﻿using System;
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
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.SlideShape;

internal sealed record SlideShapes : ISlideShapes
{
    private const long DefaultTableWidthEmu = 8128000L;
    private readonly SlidePart sdkSlidePart;

    internal SlideShapes(SlidePart sdkSlidePart)
    {
        this.sdkSlidePart = sdkSlidePart;
    }

    public int Count => this.ShapeList().Count;

    public IShape this[int index] => this.ShapeList()[index];
    
    public void Add(IShape addingShape)
    {
        var pShapeTree = this.sdkSlidePart.Slide.CommonSlideData!.ShapeTree!;
        var id = this.NextShapeId();
        var allShapeNames = this.Select(shape => shape.Name);

        if (addingShape is CopyableShape copyable)
        {
            copyable.CopyTo(id, pShapeTree, allShapeNames, this.sdkSlidePart);
        }
        else
        {
            throw new SCException($"Adding {addingShape.GetType().Name} is not supported.");
        }
    }

    public void AddAudio(int x, int y, Stream audio)
    {
        var xEmu = UnitConverter.HorizontalPixelToEmu(x);
        var yEmu = UnitConverter.VerticalPixelToEmu(y);
        var sdkPresentationDocument = (PresentationDocument)this.sdkSlidePart.OpenXmlPackage;
        var mediaDataPart = sdkPresentationDocument.CreateMediaDataPart("audio/mpeg", ".mp3");
        audio.Position = 0;
        mediaDataPart.FeedData(audio);
        var imageStream = Assembly.GetExecutingAssembly().GetStream("audio-image.png");

        var audioRef = this.sdkSlidePart.AddAudioReferenceRelationship(mediaDataPart);
        var mediaRef = this.sdkSlidePart.AddMediaReferenceRelationship(mediaDataPart);

        var audioFromFile = new A.AudioFromFile() { Link = audioRef.Id };

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
        var hyperlinkOnClick = new A.HyperlinkOnClick
            { Id = string.Empty, Action = "ppaction://media" };
        nonVisualDrawingProps.Append(hyperlinkOnClick);
        nonVisualPictureProps.Append(new P.NonVisualPictureDrawingProperties());

        var applicationNonVisualDrawingProps = nonVisualPictureProps.ApplicationNonVisualDrawingProperties!;
        applicationNonVisualDrawingProps.Append(audioFromFile);
        applicationNonVisualDrawingProps.Append(appNonVisualDrawingPropsExtensionList);
    }

    public void AddAudio(int x, int y, Stream audio, SCAudioType type)
    {
        throw new NotImplementedException();
    }

    public void AddPicture(Stream imageStream)
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
    }

    public void AddBarChart(BarChartType barChartType)
    {
        var chartFactory = new ChartGraphicFrameHandler();
        var newPGraphicFrame = chartFactory.Create(this.sdkSlidePart);

        this.sdkSlidePart.Slide.CommonSlideData!.ShapeTree!.Append(newPGraphicFrame);
    }

    public void AddVideo(int x, int y, Stream stream)
    {
        var sdkPresentationDocument = (PresentationDocument)this.sdkSlidePart.OpenXmlPackage;
        var xEmu = UnitConverter.HorizontalPixelToEmu(x);
        var yEmu = UnitConverter.VerticalPixelToEmu(y);

        var mediaDataPart = sdkPresentationDocument.CreateMediaDataPart("video/mp4", ".mp4");

        stream.Position = 0;
        mediaDataPart.FeedData(stream);
        var imgPartRId = $"rId{Guid.NewGuid().ToString().Replace("-", string.Empty).Substring(0, 5)}";
        var imagePart = this.sdkSlidePart.AddNewPart<ImagePart>("image/png", imgPartRId);
        var imageStream = Assembly.GetExecutingAssembly().GetStream("video-image.bmp");
        imagePart.FeedData(imageStream);
        var videoRr = this.sdkSlidePart.AddVideoReferenceRelationship(mediaDataPart);
        var mediaRr = this.sdkSlidePart.AddMediaReferenceRelationship(mediaDataPart);

        var pPicture = new P.Picture();

        P.NonVisualPictureProperties nonVisualPictureProperties1 = new();

        var shapeId = (uint)this.ShapeList().Max(sp => sp.Id) + 1;
        P.NonVisualDrawingProperties nonVisualDrawingProperties2 = new() { Id = shapeId, Name = $"Video{shapeId}" };
        var hyperlinkOnClick1 = new A.HyperlinkOnClick()
            { Id = string.Empty, Action = "ppaction://media" };

        A.NonVisualDrawingPropertiesExtensionList
            nonVisualDrawingPropertiesExtensionList1 = new();

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

        A.PresetGeometry presetGeometry1 = new()
            { Preset = A.ShapeTypeValues.Rectangle };
        A.AdjustValueList adjustValueList1 = new();

        presetGeometry1.Append(adjustValueList1);

        shapeProperties1.Append(transform2D1);
        shapeProperties1.Append(presetGeometry1);

        pPicture.Append(nonVisualPictureProperties1);
        pPicture.Append(blipFill1);
        pPicture.Append(shapeProperties1);

        this.sdkSlidePart.Slide.CommonSlideData!.ShapeTree!.Append(pPicture);

        DocumentFormat.OpenXml.Office2010.PowerPoint.CreationId creationId1 = new() { Val = (UInt32Value)3972997422U };
        creationId1.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");
    }

    public void AddRectangle(int x, int y, int width, int height)
    {
        var xml = new Assets(Assembly.GetExecutingAssembly()).StringOf("new-rectangle.xml");
        var sdkPShape = new P.Shape(xml);

        var position = new Position(sdkPShape); 
        position.UpdateX(x);
        position.UpdateY(y);
        
        var size = new ShapeSize(sdkPShape);
        size.UpdateWidth(width);
        size.UpdateHeight(height);

        new ShapeId(sdkPShape).Update(this.NextShapeId());
        
        this.sdkSlidePart.Slide.CommonSlideData!.ShapeTree!.Append(sdkPShape);
    }

    public void AddRoundedRectangle(int x, int y, int width, int height)
    {
        var xml = new Assets(Assembly.GetExecutingAssembly()).StringOf("new-rectangle-rounded-corners.xml");
        var sdkPShape = new P.Shape(xml);

        var position = new Position(sdkPShape); 
        position.UpdateX(x);
        position.UpdateY(y);
        
        var size = new ShapeSize(sdkPShape);
        size.UpdateWidth(width);
        size.UpdateHeight(height);

        new ShapeId(sdkPShape).Update(this.NextShapeId());
        
        this.sdkSlidePart.Slide.CommonSlideData!.ShapeTree!.Append(sdkPShape);
    }

    public void AddLine(string xml)
    {
        var newPConnectionShape = new P.ConnectionShape(xml);

        this.sdkSlidePart.Slide.CommonSlideData!.ShapeTree!.Append(newPConnectionShape);
    }

    public void AddLine(int startPointX, int startPointY, int endPointX, int endPointY)
    {
        var xml = new Assets(Assembly.GetExecutingAssembly()).StringOf("new-line.xml");
        var pConnectionShape = new P.ConnectionShape(xml);
        this.sdkSlidePart.Slide.CommonSlideData!.ShapeTree!.Append(pConnectionShape);
        
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
        
        var xEmu = UnitConverter.HorizontalPixelToEmu(x);
        var yEmu = UnitConverter.VerticalPixelToEmu(y);
        var cxEmu = UnitConverter.HorizontalPixelToEmu(cx);
        var cyEmu = UnitConverter.VerticalPixelToEmu(cy);
        var aXfrm = pConnectionShape.ShapeProperties!.Transform2D!;
        aXfrm.Offset!.X = xEmu;
        aXfrm.Offset!.Y = yEmu;
        aXfrm.Extents!.Cx = cxEmu;
        aXfrm.Extents!.Cy = cyEmu;
        aXfrm.HorizontalFlip = new BooleanValue(flipH);
        aXfrm.VerticalFlip = new BooleanValue(flipV);
    }

    public void AddTable(int xPx, int yPx, int columns, int rows)
    {
        var shapeName = this.GenerateNextTableName();
        var shapeId = this.GenerateNextShapeId();
        var xEmu = UnitConverter.HorizontalPixelToEmu(xPx);
        var yEmu = UnitConverter.VerticalPixelToEmu(yPx);
        var tableHeightEmu = SCConstants.DefaultRowHeightEmu * rows;

        var graphicFrame = new P.GraphicFrame();
        var nonVisualGraphicFrameProperties = new P.NonVisualGraphicFrameProperties();
        var nonVisualDrawingProperties = new P.NonVisualDrawingProperties { Id = (uint)shapeId, Name = shapeName };
        var nonVisualGraphicFrameDrawingProperties = new P.NonVisualGraphicFrameDrawingProperties();
        var applicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties();
        nonVisualGraphicFrameProperties.Append(nonVisualDrawingProperties);
        nonVisualGraphicFrameProperties.Append(nonVisualGraphicFrameDrawingProperties);
        nonVisualGraphicFrameProperties.Append(applicationNonVisualDrawingProperties);

        var offset = new A.Offset { X = xEmu, Y = yEmu };
        var extents = new A.Extents { Cx = DefaultTableWidthEmu, Cy = tableHeightEmu };
        var pTransform = new P.Transform(offset, extents);

        var graphic = new A.Graphic();
        var graphicData = new A.GraphicData
            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" };
        var aTable = new A.Table();

        var tableProperties = new A.TableProperties { FirstRow = true, BandRow = true };
        var tableStyleId = new A.TableStyleId
            { Text = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" };
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

        this.sdkSlidePart.Slide.CommonSlideData!.ShapeTree!.Append(graphicFrame);
    }

    public void Remove(IShape shape)
    {
        var shapes = this.ShapeList();
        if (shapes.Any(x => x != shape))
        {
            throw new SCException("Shape is not found.");
        }

        if (shape is IRemoveable removeable)
        {
            removeable.Remove();
        }
        
        throw new SCException("Shape is not cannot be removed.");
    }

    public T? GetById<T>(int shapeId)
        where T : IShape
    {
        var shape = this.ShapeList().FirstOrDefault(shape => shape.Id == shapeId);
        return (T?)shape;
    }

    public T? GetByName<T>(string shapeName)
        where T : IShape
    {
        var shape = this.GetByName(shapeName);

        return (T?)shape;
    }

    public IShape GetByName(string shapeName)
    {
        return this.ShapeList().First(shape => shape.Name == shapeName);
    }

    public IEnumerator<IShape> GetEnumerator()
    {
        return this.ShapeList().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }
    
    private (int, string) GenerateIdAndName()
    {
        var maxId = 0;
        var shapes = this.ShapeList();
        if (shapes.Any())
        {
            maxId = shapes.Max(s => s.Id);
        }

        var maxOrder = Regex.Matches(string.Join(string.Empty, shapes.Select(s => s.Name)), "\\d+",
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
        return this.ShapeList().Select(shape => shape.Id).Prepend(0).Max() + 1;
    }

    private string GenerateNextTableName()
    {
        var maxOrder = 0;
        foreach (var shape in this.ShapeList())
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

    private List<IShape> ShapeList()
    {
        var pShapeTree = this.sdkSlidePart.Slide.CommonSlideData!.ShapeTree!;
        var shapeList = new List<IShape>(pShapeTree.Count());
        foreach (var pShapeTreeElement in pShapeTree.OfType<TypedOpenXmlCompositeElement>())
        {
            if (pShapeTreeElement is P.GroupShape pGroupShape)
            {
                var groupShape = new SlideGroupShape(this.sdkSlidePart, pGroupShape);
                shapeList.Add(groupShape);
            }
            else if (pShapeTreeElement is P.ConnectionShape pConnectionShape)
            {
                var line = new SlideLine(this.sdkSlidePart, pConnectionShape);
                shapeList.Add(line);
            }
            else if (pShapeTreeElement is P.Shape pShape)
            {
                var rtSlideShape = new RootSlideAutoShape(
                    this.sdkSlidePart, 
                    pShape, 
                    new SlideAutoShape(this.sdkSlidePart, pShape));
                if (pShape.TextBody is not null)
                {
                    var textAutoShape = new TextRootSlideAutoShape(this.sdkSlidePart, rtSlideShape, pShape.TextBody);
                    shapeList.Add(textAutoShape);
                }
                else
                {
                    shapeList.Add(rtSlideShape);    
                }
            }
            else if (pShapeTreeElement is P.GraphicFrame pGraphicFrame)
            {
                var aGraphicData = pShapeTreeElement.GetFirstChild<A.Graphic>()!.GetFirstChild<A.GraphicData>();
                if (aGraphicData!.Uri!.Value!.Equals("http://schemas.openxmlformats.org/presentationml/2006/ole",
                        StringComparison.Ordinal))
                {
                    var oleObject = new SlideOLEObject(this.sdkSlidePart, pGraphicFrame);
                    shapeList.Add(oleObject);
                    continue;
                }

                var pPicture = pShapeTreeElement.Descendants<P.Picture>().FirstOrDefault();
                if (pPicture != null)
                {
                    var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
                    var blipEmbed = aBlip?.Embed;
                    if (blipEmbed is null)
                    {
                        continue;
                    }
                    
                    var picture = new SlidePicture(this.sdkSlidePart, pPicture, aBlip!);
                    shapeList.Add(picture);
                    continue;
                }
                
                if (this.IsChartPGraphicFrame(pShapeTreeElement))
                {
                    aGraphicData = pShapeTreeElement.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>() !;
                    var cChartRef = aGraphicData.GetFirstChild<C.ChartReference>() !;
                    var chartPart = (ChartPart)this.sdkSlidePart.GetPartById(cChartRef.Id!);
                    var cPlotArea = chartPart!.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea;
                    var cCharts = cPlotArea!.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
                    var cChartReference = pShapeTreeElement.GetFirstChild<A.Graphic>() !.GetFirstChild<A.GraphicData>() !
                        .GetFirstChild<C.ChartReference>() !;
                    var sdkChartPart = (ChartPart)sdkSlidePart.GetPartById(cChartReference.Id!);
                    
                    if (cCharts.Count() > 1)
                    {
                        // Combination chart

                        var combinationChart = new SlideChart(this.sdkSlidePart, (P.GraphicFrame)pShapeTreeElement, sdkChartPart);
                        shapeList.Add(combinationChart);
                        continue;
                    }

                    var chartTypeName = cCharts.Single().LocalName;

                    if (chartTypeName == "lineChart")
                    {
                        var lineChart = new SlideChart(this.sdkSlidePart, (P.GraphicFrame)pShapeTreeElement,
                            sdkChartPart);
                        shapeList.Add(lineChart);
                        continue;
                    }

                    if (chartTypeName == "barChart")
                    {
                        var barChart = new SlideChart(this.sdkSlidePart, (P.GraphicFrame)pShapeTreeElement, sdkChartPart);
                        shapeList.Add(barChart);
                        continue;
                    }

                    if (chartTypeName == "pieChart")
                    {
                        var pieChart = new SlideChart(this.sdkSlidePart, (P.GraphicFrame)pShapeTreeElement, sdkChartPart);
                        shapeList.Add(pieChart);
                        continue;
                    }

                    if (chartTypeName == "scatterChart")
                    {
                        var scatterChart = new SlideChart(this.sdkSlidePart, (P.GraphicFrame)pShapeTreeElement,
                            sdkChartPart);
                        shapeList.Add(scatterChart);
                        continue;
                    }

                    var chart = new SlideChart(this.sdkSlidePart,(P.GraphicFrame)pShapeTreeElement, sdkChartPart);
                    shapeList.Add(chart);
                }
                else if (this.IsTablePGraphicFrame(pShapeTreeElement))
                {
                    var table = new SlideTable(this.sdkSlidePart, pShapeTreeElement);
                    shapeList.Add(table);
                }
                
            }
            else if (pShapeTreeElement is P.Picture pPicture)
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
                            var mediaShape = new SlideMediaShape(this.sdkSlidePart, pPicture);
                            shapeList.Add(mediaShape);
                        }

                        continue;
                    }
                    case A.VideoFromFile:
                    {
                        var mediaShape = new SlideMediaShape(this.sdkSlidePart, pPicture);
                        shapeList.Add(mediaShape);
                        continue;
                    }
                }

                var aBlip = pPicture.GetFirstChild<P.BlipFill>()?.Blip;
                var blipEmbed = aBlip?.Embed;
                if (blipEmbed is null)
                {
                    continue;
                }

                var picture = new SlidePicture(this.sdkSlidePart, pPicture, aBlip!);
                shapeList.Add(picture);
            }
        }

        return shapeList;
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

    private P.Picture CreatePPicture(Stream imageStream, string shapeName)
    {
        var imgPartRId = this.sdkSlidePart.GetNextRelationshipId();
        var imagePart = this.sdkSlidePart.AddNewPart<ImagePart>("image/png", imgPartRId);
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
        var blip = new A.Blip { Embed = imgPartRId };
        var stretch = new A.Stretch();
        blipFill.Append(blip);
        blipFill.Append(stretch);

        var transform2D = new A.Transform2D(
            new A.Offset { X = 0, Y = 0 },
            new A.Extents { Cx = 0, Cy = 0 });

        var presetGeometry = new A.PresetGeometry
            { Preset = A.ShapeTypeValues.Rectangle };
        var shapeProperties = new P.ShapeProperties();
        shapeProperties.Append(transform2D);
        shapeProperties.Append(presetGeometry);

        var pPicture = new P.Picture();
        pPicture.Append(nonVisualPictureProperties);
        pPicture.Append(blipFill);
        pPicture.Append(shapeProperties);

        this.sdkSlidePart.Slide.CommonSlideData!.ShapeTree!.Append(pPicture);

        return pPicture;
    }
    
    private int NextShapeId()
    {
        var shapes = this.ShapeList();
        if (shapes.Any())
        {
            return shapes.Select(shape => shape.Id).Prepend(0).Max() + 1;
        }

        return 1;
    }
}