// ReSharper disable InconsistentNaming
// ReSharper disable UseObjectOrCollectionInitializer

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Assets;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Presentations;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tables;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed class SlideShapeCollection : ISlideShapeCollection
{
    private readonly SlidePart slidePart;
    private readonly IShapeCollection shapes;
    private readonly MediaShapeCollection mediaShapes;

    internal SlideShapeCollection(IShapeCollection shapes, SlidePart slidePart, MediaCollection mediaCollection)
    {
        this.shapes = shapes;
        this.mediaShapes = new MediaShapeCollection(shapes, slidePart, mediaCollection);
        this.slidePart = slidePart;
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

    public void AddAudio(int x, int y, Stream audio) => this.mediaShapes.AddAudio(x, y, audio);

    public void AddAudio(int x, int y, Stream audio, AudioType type) => this.mediaShapes.AddAudio(x, y, audio, type);

    public void AddPicture(Stream imageStream) => this.mediaShapes.AddPicture(imageStream);

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

    public void AddBarChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName)
    {
        new SCSlidePart(this.slidePart).AddBarChart(x, y, width, height, categoryValues, seriesName);
    }

    /// <summary>
    ///     Adds a SmartArt graphic to the slide.
    /// </summary>
    /// <param name="x">The x-coordinate of the SmartArt graphic.</param>
    /// <param name="y">The y-coordinate of the SmartArt graphic.</param>
    /// <param name="width">The width of the SmartArt graphic.</param>
    /// <param name="height">The height of the SmartArt graphic.</param>
    /// <param name="smartArtType">The type of SmartArt graphic to add.</param>
    /// <returns>The added SmartArt graphic.</returns>
    public ISmartArt AddSmartArt(int x, int y, int width, int height, SmartArtType smartArtType)
    {
        return new SCSlidePart(this.slidePart).AddSmartArt(x, y, width, height, smartArtType);
    }

    public IGroup Group(IShape[] groupingShapes)
    {
        // Create a new group shape
        var groupShape = new P.GroupShape();
        
        // Create non-visual properties for the group shape
        var nonVisualGroupShapeProperties = new P.NonVisualGroupShapeProperties();
        var idAndName = this.GenerateIdAndName();
        var nonVisualDrawingProperties = new P.NonVisualDrawingProperties
        {
            Id = (uint)idAndName.Item1,
            Name = idAndName.Item2
        };
        var nonVisualGroupShapeDrawingProperties = new P.NonVisualGroupShapeDrawingProperties();
        var applicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties();
        
        nonVisualGroupShapeProperties.Append(nonVisualDrawingProperties);
        nonVisualGroupShapeProperties.Append(nonVisualGroupShapeDrawingProperties);
        nonVisualGroupShapeProperties.Append(applicationNonVisualDrawingProperties);
        
        // Create group shape properties
        var groupShapeProperties = new P.GroupShapeProperties();
        
        // Calculate the bounding box for all shapes
        decimal minX = decimal.MaxValue;
        decimal minY = decimal.MaxValue;
        decimal maxX = decimal.MinValue;
        decimal maxY = decimal.MinValue;
        
        foreach (var shape in groupingShapes)
        {
            minX = Math.Min(minX, shape.X);
            minY = Math.Min(minY, shape.Y);
            maxX = Math.Max(maxX, shape.X + shape.Width);
            maxY = Math.Max(maxY, shape.Y + shape.Height);
        }
        
        // Create transform group
        var transformGroup = new A.TransformGroup();
        var offset = new A.Offset 
        { 
            X = (int)minX, 
            Y = (int)minY 
        };
        var extents = new A.Extents 
        { 
            Cx = (int)(maxX - minX), 
            Cy = (int)(maxY - minY) 
        };
        var childOffset = new A.ChildOffset { X = 0, Y = 0 };
        var childExtents = new A.ChildExtents { Cx = extents.Cx, Cy = extents.Cy };
        
        transformGroup.Append(offset);
        transformGroup.Append(extents);
        transformGroup.Append(childOffset);
        transformGroup.Append(childExtents);
        
        groupShapeProperties.Append(transformGroup);
        
        // Add properties to the group shape
        groupShape.Append(nonVisualGroupShapeProperties);
        groupShape.Append(groupShapeProperties);
        
        // Move each shape to the group
        foreach (var shape in groupingShapes)
        {
            // Get the OpenXml element for the shape
            var openXmlElement = shape.SDKOpenXmlElement;
            
            // Remove the shape from its current parent
            if (openXmlElement.Parent is not null)
            {
                openXmlElement.Remove();
            }
            
            // Add the shape to the group
            groupShape.Append(openXmlElement);
        }
        
        // Add the group shape to the slide
        this.slidePart.Slide.CommonSlideData!.ShapeTree!.Append(groupShape);
        
        // Create and return the group object
        var shapeObj = new Shape(groupShape);

        foreach (var grouping in groupingShapes)
        {
            grouping.Remove();
        }
        
        return new Group(shapeObj, groupShape);
    }

    public void AddVideo(int x, int y, Stream stream)
    {
        var presDocument = (PresentationDocument)this.slidePart.OpenXmlPackage;
        var mediaDataPart = presDocument.CreateMediaDataPart("video/mp4", ".mp4");
        mediaDataPart.FeedData(stream);
        var imagePartRId = $"rId{Guid.NewGuid().ToString().Replace("-", string.Empty)[..5]}";
        var imagePart = this.slidePart.AddNewPart<ImagePart>("image/png", imagePartRId);
        var imageStream = new AssetCollection(Assembly.GetExecutingAssembly()).StreamOf("video image.bmp");
        imagePart.FeedData(imageStream);
        var videoRr = this.slidePart.AddVideoReferenceRelationship(mediaDataPart);
        var mediaRr = this.slidePart.AddMediaReferenceRelationship(mediaDataPart);

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

        this.slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pPicture);

        DocumentFormat.OpenXml.Office2010.PowerPoint.CreationId creationId = new() { Val = (UInt32Value)3972997422U };
        creationId.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");
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

    public void AddShape(int x, int y, int width, int height, Geometry geometry, string text)
    {
        // First add the basic shape
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
        addedShape.TextBox!.SetText(text);
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
        this.AddTable(x, y, columnsCount, rowsCount, CommonTableStyles.MediumStyle2Accent1);

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

        const long DefaultTableWidthEMUs = 8128000L;
        var offset = new A.Offset { X = xEmu, Y = yEmu };
        var extents = new A.Extents { Cx = DefaultTableWidthEMUs, Cy = tableHeightEmu };
        var pTransform = new P.Transform(offset, extents);

        var graphic = new A.Graphic();
        var graphicData = new A.GraphicData { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" };
        var aTable = new A.Table();

        var tableProperties = new A.TableProperties { FirstRow = true, BandRow = true };
        var tableStyleId = new A.TableStyleId { Text = ((TableStyle)style).Guid };
        tableProperties.Append(tableStyleId);

        var tableGrid = new A.TableGrid();
        var gridWidthEmu = DefaultTableWidthEMUs / columnsCount;
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

    public IShape GetById(int id) => this.GetById<IShape>(id);

    public T GetById<T>(int id)
        where T : IShape => this.shapes.GetById<T>(id);

    public T GetByName<T>(string name)
        where T : IShape => this.shapes.Shape<T>(name);

    public T Shape<T>(string name)
        where T : IShape => this.shapes.Shape<T>(name);

    public IShape Shape(string name) => this.shapes.Shape(name);

    public T Last<T>()
        where T : IShape => this.shapes.Last<T>();

    public IEnumerator<IShape> GetEnumerator() => this.shapes.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

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
}