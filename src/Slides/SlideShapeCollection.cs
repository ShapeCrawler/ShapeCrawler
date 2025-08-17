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
using ShapeCrawler.Extensions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tables;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed class SlideShapeCollection(ISlideShapeCollection shapes, SlidePart slidePart) : ISlideShapeCollection
{
    public int Count => shapes.Count;

    public IShape this[int index] => shapes[index];

    public void Add(IShape addingShape)
    {
        var pShapeTree = slidePart.Slide.CommonSlideData!.ShapeTree!;
        switch (addingShape)
        {
            case PictureShape picture:
                picture.CopyTo(pShapeTree);
                break;
            case TextShape textShape:
                textShape.CopyTo(pShapeTree);
                break;
            case TableShape table:
                table.CopyTo(pShapeTree);
                break;
            case Shape shape:
                shape.CopyTo(pShapeTree);
                break;
            default:
                throw new SCException("Unsupported shape type for adding.");
        }
    }

    public void AddAudio(int x, int y, Stream audio) => shapes.AddAudio(x, y, audio);

    public void AddAudio(int x, int y, Stream audio, AudioType type) => shapes.AddAudio(x, y, audio, type);

    public void AddVideo(int x, int y, Stream stream) => shapes.AddVideo(x, y, stream);

    public void AddPicture(Stream imageStream) => shapes.AddPicture(imageStream);

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
        SmartArtType smartArtType)
        => new SCSlidePart(slidePart).AddSmartArt(x, y, width, height, smartArtType);

    public IShape Group(IShape[] groupingShapes)
    {
        var nonVisualGroupShapeProperties = new P.NonVisualGroupShapeProperties();
        var idAndName = this.GenerateIdAndName();
        var nonVisualDrawingProperties = new P.NonVisualDrawingProperties
        {
            Id = (uint)idAndName.Item1, Name = idAndName.Item2
        };
        var nonVisualGroupShapeDrawingProperties = new P.NonVisualGroupShapeDrawingProperties();
        var applicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties();

        nonVisualGroupShapeProperties.Append(nonVisualDrawingProperties);
        nonVisualGroupShapeProperties.Append(nonVisualGroupShapeDrawingProperties);
        nonVisualGroupShapeProperties.Append(applicationNonVisualDrawingProperties);

        var groupShapeProperties = new P.GroupShapeProperties();

        decimal minX = decimal.MaxValue;
        decimal minY = decimal.MaxValue;
        decimal maxX = decimal.MinValue;
        decimal maxY = decimal.MinValue;

        foreach (var groupingShape in groupingShapes)
        {
            minX = Math.Min(minX, groupingShape.X);
            minY = Math.Min(minY, groupingShape.Y);
            maxX = Math.Max(maxX, groupingShape.X + groupingShape.Width);
            maxY = Math.Max(maxY, groupingShape.Y + groupingShape.Height);
        }

        var transformGroup = new A.TransformGroup();
        var offset = new A.Offset { X = (int)minX, Y = (int)minY };
        var extents = new A.Extents { Cx = (int)(maxX - minX), Cy = (int)(maxY - minY) };
        var childOffset = new A.ChildOffset { X = 0, Y = 0 };
        var childExtents = new A.ChildExtents { Cx = extents.Cx, Cy = extents.Cy };

        transformGroup.Append(offset);
        transformGroup.Append(extents);
        transformGroup.Append(childOffset);
        transformGroup.Append(childExtents);

        groupShapeProperties.Append(transformGroup);

        var pGroupShape = new P.GroupShape();
        pGroupShape.Append(nonVisualGroupShapeProperties);
        pGroupShape.Append(groupShapeProperties);

        foreach (var groupingShape in groupingShapes)
        {
            // Get the OpenXml element for the shape
            var openXmlElement = groupingShape.SDKOpenXmlElement;

            // Remove the shape from its current parent
            if (openXmlElement.Parent is not null)
            {
                openXmlElement.Remove();
            }
            
            pGroupShape.Append(openXmlElement);
        }

        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pGroupShape);

        foreach (var grouping in groupingShapes)
        {
            grouping.Remove();
        }

        return new GroupShape(pGroupShape);
    }

    public void AddShape(int x, int y, int width, int height, Geometry geometry = Geometry.Rectangle)
    {
        var xml = new AssetCollection(Assembly.GetExecutingAssembly()).StringOf("new rectangle.xml");
        var pShape = new P.Shape(xml);
        var nextShapeId = this.GetNextShapeId();
        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pShape);

        var addedShape = shapes.Last<TextShape>();
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
        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pShape);

        var addedShape = shapes.Last<TextShape>();
        addedShape.Name = geometry.ToString();
        addedShape.X = x;
        addedShape.Y = y;
        addedShape.Width = width;
        addedShape.Height = height;
        addedShape.Id = nextShapeId;
        addedShape.GeometryType = geometry;
        addedShape.TextBox.SetText(text);
    }

    public void AddLine(string xml)
    {
        var newPConnectionShape = new P.ConnectionShape(xml);

        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(newPConnectionShape);
    }

    public void AddLine(int startPointX, int startPointY, int endPointX, int endPointY)
    {
        var xml = new AssetCollection(Assembly.GetExecutingAssembly()).StringOf("new line.xml");
        var pConnectionShape = new P.ConnectionShape(xml);
        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pConnectionShape);

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

    public void AddTable(int x, int y, int columnsCount, int rowsCount)
        => this.AddTable(x, y, columnsCount, rowsCount, CommonTableStyles.MediumStyle2Accent1);

    public void AddTable(int x, int y, int columnsCount, int rowsCount, ITableStyle style)
    {
        var shapeName = this.GenerateNextTableName();
        var xEmu = new Points(x).AsEmus();
        var yEmu = new Points(y).AsEmus();
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

        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(graphicFrame);
    }

    public IShape GetById(int id) => shapes.GetById(id);

    public T GetById<T>(int id)
        where T : IShape => shapes.GetById<T>(id);

    public T GetByName<T>(string name)
        where T : IShape => shapes.Shape<T>(name);

    public T Shape<T>(string name)
        where T : IShape => shapes.Shape<T>(name);

    public IShape Shape(string name) => shapes.Shape(name);

    public T Last<T>()
        where T : IShape => shapes.Last<T>();

    public IEnumerator<IShape> GetEnumerator() => shapes.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private int GetNextShapeId()
    {
        if (shapes.Any())
        {
            return shapes.Select(shape => shape.Id).Prepend(0).Max() + 1;
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
        foreach (var shape in shapes)
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