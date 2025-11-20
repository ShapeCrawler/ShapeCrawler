// ReSharper disable InconsistentNaming
// ReSharper disable UseObjectOrCollectionInitializer

using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Assets;
using ShapeCrawler.Groups;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tables;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed class SlideShapeCollection(ISlideShapeCollection shapes, SlidePart slidePart) : ISlideShapeCollection
{
    private readonly NewShapeProperties newShapeProperties = new(shapes);
    private readonly PlaceholderShapes placeholderShape = new(shapes, slidePart);
    private readonly ConnectionShape connectionShape = new(slidePart, new NewShapeProperties(shapes));

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
        SmartArtType smartArtType)
        => new SCSlidePart(slidePart).AddSmartArt(x, y, width, height, smartArtType);

    public IShape Group(IShape[] groupingShapes) => new GroupShape(new P.GroupShape(), groupingShapes, this.newShapeProperties, slidePart);

    public void AddShape(int x, int y, int width, int height, Geometry geometry = Geometry.Rectangle)
    {
        var xml = new AssetCollection(Assembly.GetExecutingAssembly()).StringOf("new rectangle.xml");
        var pShape = new P.Shape(xml);
        var nextShapeId = this.newShapeProperties.Id();
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
        var nextShapeId = this.newShapeProperties.Id();
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
        => this.connectionShape.Create(startPointX, startPointY, endPointX, endPointY);

    public void AddTable(int x, int y, int columnsCount, int rowsCount)
        => this.AddTable(x, y, columnsCount, rowsCount, CommonTableStyles.MediumStyle2Accent1);

    public void AddTable(int x, int y, int columnsCount, int rowsCount, ITableStyle style)
    {
        var shapeName = this.newShapeProperties.TableName();
        var xEmu = new Points(x).AsEmus();
        var yEmu = new Points(y).AsEmus();
        var tableHeightEmu = Constants.DefaultRowHeightEmu * rowsCount;

        var graphicFrame = new P.GraphicFrame();
        var nonVisualGraphicFrameProperties = new P.NonVisualGraphicFrameProperties();
        var nonVisualDrawingProperties = new P.NonVisualDrawingProperties
        {
            Id = (uint)this.newShapeProperties.Id(), Name = shapeName
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
            var aTableRow = new A.TableRow { Height = Constants.DefaultRowHeightEmu };
            for (var i2 = 0; i2 < columnsCount; i2++)
            {
                new SCATableRow(aTableRow).AddNewCell();
            }
        
            aTable.Append(aTableRow);
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

    public IShape AddDateAndTime() => this.placeholderShape.AddDateAndTime();

    public IShape AddFooter() => this.placeholderShape.AddFooter();

    public IShape AddSlideNumber() => this.placeholderShape.AddSlideNumber();
}